"""
api/transform.py — Vercel Serverless Function (Python)
POST /api/transform  →  multipart/form-data { file: Salesforce.xlsx }
Returns the transformed Excel file as a download.
"""

import os
import sys
import json
import tempfile
import traceback

# ── Path setup ──────────────────────────────────────────────────────────────
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from transfo_odoo import transform_import_odoo

# Référentiel intégré au projet
REFERENTIEL = os.path.join(ROOT, "data", "Comptes_analytiques.xlsx")


# ── WSGI app — utilisé par Vercel ────────────────────────────────────────────

def app(environ, start_response):
    method = environ.get("REQUEST_METHOD", "GET").upper()

    # CORS preflight
    if method == "OPTIONS":
        start_response("200 OK", _cors_headers([
            ("Content-Length", "0"),
            ("Content-Type", "text/plain"),
        ]))
        return [b""]

    if method != "POST":
        return _error(start_response, 405, "Méthode non autorisée. Utilisez POST.")

    try:
        # ── Lire le corps de la requête ──────────────────────────────────
        content_length = int(environ.get("CONTENT_LENGTH") or 0)
        if content_length == 0:
            return _error(start_response, 400, "Corps de requête vide.")

        body = environ["wsgi.input"].read(content_length)

        # ── Parser le multipart ──────────────────────────────────────────
        content_type = environ.get("CONTENT_TYPE", "")
        file_bytes = _extract_file_from_multipart(body, content_type)

        if file_bytes is None:
            return _error(start_response, 400,
                          "Champ 'file' introuvable dans le formulaire multipart.")

        if len(file_bytes) == 0:
            return _error(start_response, 400, "Le fichier reçu est vide.")

        # ── Vérifier le référentiel ──────────────────────────────────────
        if not os.path.exists(REFERENTIEL):
            return _error(start_response, 500,
                          "Référentiel introuvable sur le serveur. Contactez l'administrateur.")

        # ── Écrire le fichier entrant dans un temp ───────────────────────
        tmp_in = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        tmp_in.write(file_bytes)
        tmp_in.close()
        tmp_in_path = tmp_in.name
        tmp_out_path = tmp_in_path.replace(".xlsx", "_transforme.xlsx")

        # ── Transformation ───────────────────────────────────────────────
        try:
            df_result, _ = transform_import_odoo(
                input_excel_path=tmp_in_path,
                referentiel_path=REFERENTIEL,
                output_excel_path=tmp_out_path,
                sheet_name="Import Odoo",
            )
        except Exception as e:
            print(f"[transfo_error] {e}\n{traceback.format_exc()}")
            return _error(start_response, 500, f"Erreur de transformation : {e}")
        finally:
            _safe_delete(tmp_in_path)

        # ── Lire le résultat ─────────────────────────────────────────────
        try:
            with open(tmp_out_path, "rb") as f:
                result_bytes = f.read()
        finally:
            _safe_delete(tmp_out_path)

        rows = str(len(df_result))
        cols = str(len(df_result.columns))
        filename = "Salesforce_transforme.xlsx"

        headers = _cors_headers([
            ("Content-Type",
             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
            ("Content-Disposition", f'attachment; filename="{filename}"'),
            ("Content-Length", str(len(result_bytes))),
            ("X-Rows", rows),
            ("X-Cols", cols),
            ("X-Filename", filename),
        ])
        start_response("200 OK", headers)
        return [result_bytes]

    except Exception as e:
        print(f"[unhandled_error] {e}\n{traceback.format_exc()}")
        return _error(start_response, 500, f"Erreur interne : {e}")


# ── Helpers ──────────────────────────────────────────────────────────────────

def _cors_headers(extra=None):
    base = [
        ("Access-Control-Allow-Origin", "*"),
        ("Access-Control-Allow-Methods", "POST, OPTIONS"),
        ("Access-Control-Allow-Headers", "Content-Type"),
        ("Access-Control-Expose-Headers", "X-Rows, X-Cols, X-Filename"),
    ]
    if extra:
        base.extend(extra)
    return base


def _error(start_response, code, message):
    body = json.dumps({"detail": message}).encode("utf-8")
    status = {
        400: "400 Bad Request",
        404: "404 Not Found",
        405: "405 Method Not Allowed",
        500: "500 Internal Server Error",
    }.get(code, f"{code} Error")
    start_response(status, _cors_headers([
        ("Content-Type", "application/json"),
        ("Content-Length", str(len(body))),
    ]))
    return [body]


def _safe_delete(path):
    try:
        if path and os.path.exists(path):
            os.unlink(path)
    except Exception:
        pass


def _extract_file_from_multipart(body: bytes, content_type: str):
    """Parse manuellement le multipart/form-data (sans cgi déprecié)."""
    boundary = None
    for part in content_type.split(";"):
        part = part.strip()
        if part.startswith("boundary="):
            boundary = part[len("boundary="):].strip().strip('"')
            break

    if not boundary:
        return None

    boundary_bytes = ("--" + boundary).encode()
    parts = body.split(boundary_bytes)

    for part in parts:
        if b'name="file"' not in part and b"name='file'" not in part:
            continue
        if b"\r\n\r\n" in part:
            _, content = part.split(b"\r\n\r\n", 1)
        elif b"\n\n" in part:
            _, content = part.split(b"\n\n", 1)
        else:
            continue
        content = content.rstrip(b"\r\n--")
        return content

    return None
