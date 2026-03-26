"""
api/transform.py — Vercel Serverless Function
Endpoint POST /api/transform
Reçoit un fichier Excel (Salesforce.xlsx), applique la transformation Odoo,
retourne le fichier transformé en téléchargement direct.
"""

from http.server import BaseHTTPRequestHandler
import json
import os
import sys
import io
import cgi
import tempfile

# Ajouter le répertoire parent au path pour importer transfo_odoo
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..'))

from transfo_odoo import transform_import_odoo


class handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(200)
        self._set_cors_headers()
        self.end_headers()

    def do_POST(self):
        if self.path != '/api/transform':
            self.send_error(404, 'Not Found')
            return

        try:
            # Parse multipart form data
            content_type = self.headers.get('Content-Type', '')
            if 'multipart/form-data' not in content_type:
                self._send_json_error(400, 'Content-Type multipart/form-data requis')
                return

            # Read body
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)

            # Parse using cgi module
            environ = {
                'REQUEST_METHOD': 'POST',
                'CONTENT_TYPE': content_type,
                'CONTENT_LENGTH': str(content_length),
            }
            form = cgi.FieldStorage(
                fp=io.BytesIO(body),
                headers=self.headers,
                environ=environ
            )

            if 'file' not in form:
                self._send_json_error(400, "Champ 'file' manquant dans le formulaire")
                return

            file_item = form['file']
            file_data = file_item.file.read()

            if not file_data:
                self._send_json_error(400, 'Fichier vide reçu')
                return

            # Write input to temp file
            with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as tmp_in:
                tmp_in.write(file_data)
                tmp_in_path = tmp_in.name

            tmp_out_path = tmp_in_path.replace('.xlsx', '_transforme.xlsx')

            # Référentiel : intégré dans le projet
            referentiel_path = os.path.join(
                os.path.dirname(__file__), '..', 'data', 'Comptes_analytiques.xlsx'
            )

            if not os.path.exists(referentiel_path):
                self._send_json_error(500, f'Référentiel introuvable : {referentiel_path}')
                return

            # Transformation
            df_result, out_path = transform_import_odoo(
                input_excel_path=tmp_in_path,
                referentiel_path=referentiel_path,
                output_excel_path=tmp_out_path,
                sheet_name='Import Odoo'
            )

            # Read result
            with open(tmp_out_path, 'rb') as f:
                result_bytes = f.read()

            # Cleanup
            os.unlink(tmp_in_path)
            os.unlink(tmp_out_path)

            rows = len(df_result)
            cols = len(df_result.columns)
            output_filename = 'Salesforce_transforme.xlsx'

            # Send response
            self.send_response(200)
            self._set_cors_headers()
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', f'attachment; filename="{output_filename}"')
            self.send_header('Content-Length', str(len(result_bytes)))
            self.send_header('X-Rows', str(rows))
            self.send_header('X-Cols', str(cols))
            self.send_header('X-Filename', output_filename)
            self.end_headers()
            self.wfile.write(result_bytes)

        except Exception as e:
            import traceback
            tb = traceback.format_exc()
            print(f"[ERROR] {e}\n{tb}")
            # Cleanup on error
            for p in [tmp_in_path if 'tmp_in_path' in dir() else None,
                      tmp_out_path if 'tmp_out_path' in dir() else None]:
                if p and os.path.exists(p):
                    try:
                        os.unlink(p)
                    except Exception:
                        pass
            self._send_json_error(500, str(e))

    def _set_cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Access-Control-Expose-Headers', 'X-Rows, X-Cols, X-Filename')

    def _send_json_error(self, code, message):
        body = json.dumps({'detail': message}).encode('utf-8')
        self.send_response(code)
        self._set_cors_headers()
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)
