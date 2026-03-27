
# -*- coding: utf-8 -*-
import json
import unicodedata
from pathlib import Path

import numpy as np
import pandas as pd


# =========================
# OUTILS GÉNÉRAUX
# =========================
def normalize_text(value):
    if pd.isna(value):
        return ""
    value = str(value).strip()
    value = unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii")
    value = value.upper()
    value = " ".join(value.split())
    return value


def to_str(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def safe_round(value, decimals=3):
    if pd.isna(value) or value == "":
        return np.nan
    try:
        return round(float(str(value).replace(",", ".")), decimals)
    except Exception:
        return np.nan


def format_number_fr(value, decimals=3):
    if pd.isna(value) or value == "":
        return ""
    try:
        value = round(float(value), decimals)
        txt = f"{value:.{decimals}f}"
        txt = txt.rstrip("0").rstrip(".")
        return txt.replace(".", ",")
    except Exception:
        return to_str(value)


def parse_percent_to_int(value, default=0):
    """
    Convertit proprement:
    - 0.75 -> 75
    - 75 -> 75
    - "" / NaN -> 0
    """
    if pd.isna(value) or value == "":
        return int(default)
    try:
        txt = str(value).strip().replace("%", "").replace(",", ".")
        if txt == "":
            return int(default)
        num = float(txt)
        if np.isnan(num):
            return int(default)
        if 0 <= num <= 1:
            num *= 100
        return int(round(num))
    except Exception:
        return int(default)


def clean_text_series(series):
    return series.fillna("").astype(str).str.strip()


# =========================
# RÉFÉRENTIEL
# =========================
ALIASES_02 = {
    "CONQUETE": "CONQUETE (AUTRES)",
    "CONQUÊTE": "CONQUETE (AUTRES)",
    "RENOUVELLEMENT": "RENOUVELLEMENT",
    "VENTE ADDITIONNELLE": "VENTE ADDITIONNELLE",
}

ALIASES_03 = {
    "ELECTRICITE": "ELECTRICITE",
    "ÉLECTRICITÉ": "ELECTRICITE",
    "GAZ": "GAZ",
}

ALIASES_04 = {
    "ENERGY PRO CONSULTING": "ENERGY PRO",
    "CABINET BRUERE ENERGIES": "CABINET BRUERE ENERGIES (CBE)",
    "REZO ENERGY": "REZO ENERGY - NES SALES",
    "STEPHANIE MIROUX": "STEPHANIE MIROUX",
    "STÉPHANIE MIROUX": "STEPHANIE MIROUX",
    "NOLWENN FAVRE": "NOLWENN FAVRE",
    "BENOIT VILCOT": "BENOIT VILCOT",
    "BENOIT VILCOT ": "BENOIT VILCOT",
    "MYLENE PROST": "MYLENE PROST",
    "MYLÈNE PROST": "MYLENE PROST",
    "PIERRE-JEAN HAURE": "PIERRE-JEAN HAURE",
    "ENOPTEA": "ENOPTEA",
    "VS ENERGIE": "VS ENERGIE",
    "D.A. CONSULTING": "D.A. CONSULTING",
}


def load_referentiel(referentiel_path):
    ref = pd.read_excel(referentiel_path, sheet_name=0, dtype=object)
    ref.columns = [str(c).strip() for c in ref.columns]
    required_cols = {"Plan", "Compte analytique", "ID"}
    missing = required_cols - set(ref.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes dans le référentiel : {missing}")

    ref["Plan_norm"] = ref["Plan"].apply(normalize_text)
    ref["Compte_norm"] = ref["Compte analytique"].apply(normalize_text)
    ref["ID"] = ref["ID"].astype(str).str.strip()

    maps = {}
    for plan_name in ref["Plan_norm"].dropna().unique():
        subset = ref[ref["Plan_norm"] == plan_name].copy()
        maps[plan_name] = dict(zip(subset["Compte_norm"], subset["ID"]))
    return maps, ref


def get_ref_id(maps, plan, label, default=None, aliases=None):
    plan_norm = normalize_text(plan)
    label_norm = normalize_text(label)
    if aliases and label_norm in aliases:
        label_norm = normalize_text(aliases[label_norm])
    return maps.get(plan_norm, {}).get(label_norm, default)


def build_codes_from_referentiel(df, maps):
    # BU
    bu_externe = get_ref_id(maps, "01 - BU", "01- Courtage Externe", default="49")
    bu_interne = get_ref_id(maps, "01 - BU", "02 - Courtage Internes", default="48")

    df["Gestionnaire-BU"] = np.where(
        clean_text_series(df["Apporteur d'affaire"]).str.upper() != "CAPITOLE ENERGIE",
        str(bu_externe),
        str(bu_interne),
    )

    # Type
    df["TypeAnalytique"] = df["Type"].apply(
        lambda x: get_ref_id(maps, "02 - Niveau", x, default="", aliases=ALIASES_02)
    )

    # Produit
    df["ProduitAnalytique"] = df["Lignes de la commande/Produit"].apply(
        lambda x: get_ref_id(maps, "03 - Niveau", x, default="", aliases=ALIASES_03)
    )

    # Niveau 4 = apporteur si trouvé, sinon vendeur, sinon 0
    def map_codes_ok(row):
        apporteur = row.get("Apporteur d'affaire", "")
        vendeur = row.get("Vendeur", "")
        code_apporteur = get_ref_id(maps, "04 - Niveau", apporteur, default=None, aliases=ALIASES_04)
        if code_apporteur is not None:
            return str(code_apporteur)

        code_vendeur = get_ref_id(maps, "04 - Niveau", vendeur, default=None, aliases=ALIASES_04)
        if code_vendeur is not None:
            return str(code_vendeur)

        return "0"

    df["Codes OK"] = df.apply(map_codes_ok, axis=1)

    ce_aa = get_ref_id(maps, "04 - Niveau", "Capitole Energie AA", default="76", aliases=ALIASES_04)
    df["Capitole Energie AA"] = np.where(
        df["Gestionnaire-BU"].astype(str).str.startswith("49"),
        str(ce_aa),
        "",
    )
    return df


# =========================
# MISE EN FORME
# =========================
def format_date_for_note(value):
    if pd.isna(value) or value == "":
        return ""
    try:
        dt = pd.to_datetime(value, errors="coerce", dayfirst=True)
        if pd.notna(dt):
            return dt.strftime("%d/%m/%Y")
    except Exception:
        pass
    return str(value).strip()


def build_distribution_json(row):
    bu = to_str(row.get("Gestionnaire-BU"))
    type_ = to_str(row.get("TypeAnalytique"))
    produit = to_str(row.get("ProduitAnalytique"))
    code = to_str(row.get("Codes OK"))

    pct_main = parse_percent_to_int(row.get("Pourcentage.1"), default=0)
    payload = {}

    key_main = ",".join([bu, type_, produit, code])
    payload[key_main] = pct_main

    ce = to_str(row.get("Capitole Energie AA"))
    if ce:
        pct_ce = max(0, 100 - pct_main)
        key_ce = ",".join([bu, type_, produit, ce])
        payload[key_ce] = pct_ce

    return json.dumps(payload, ensure_ascii=False)


def build_product_row(row):
    energie = to_str(row.get("Lignes de la commande/Produit"))
    return {
        "Client": to_str(row.get("Client")),
        "Lignes de la commande/Produit": energie,
        "Lignes de la commande/Type d'affichage": "",
        "Lignes de la commande/Description1.1": f"Contrat Energie {energie}".strip(),
        "Lignes de la commande/Description2": "",
        "Lignes de la commande/Description": "",
        "Lignes de la commande/Description3": "",
        "Lignes de la commande/Description4": "",
        "Lignes de la commande/Description5": "",
        "Lignes de la commande/Description6": "",
        "Lignes de la commande/Quantité": "1",
        "Lignes de la commande/Prix Unitaire": row.get("Lignes de la commande/Prix Unitaire", 0) if not pd.isna(row.get("Lignes de la commande/Prix Unitaire", 0)) else 0,
        "Lignes de la commande/Distribution Analytique": build_distribution_json(row),
    }


def build_note_row(row):
    raison_sociale = to_str(row.get("Lignes de la commande/Description1"))
    pdl = to_str(row.get("PDL"))
    date_signature = format_date_for_note(row.get("Lignes de la commande/Date de signature"))
    date_debut = format_date_for_note(row.get("Lignes de la commande/Description3"))
    date_fin = format_date_for_note(row.get("Lignes de la commande/Description4"))
    duree = to_str(row.get("Lignes de la commande/Description5"))
    car = format_number_fr(safe_round(row.get("Lignes de la commande/Description6"), 3), 3)

    return {
        "Client": "",
        "Lignes de la commande/Produit": "",
        "Lignes de la commande/Type d'affichage": "NOTE",
        "Lignes de la commande/Description1.1": f"RAISON SOCIALE : {raison_sociale}" if raison_sociale else "RAISON SOCIALE : ",
        "Lignes de la commande/Description2": f"PDL : {pdl}" if pdl else "PDL : ",
        "Lignes de la commande/Description": f"Date de signature : {date_signature}" if date_signature else "Date de signature : ",
        "Lignes de la commande/Description3": f"Date de début contrat : {date_debut}" if date_debut else "Date de début contrat : ",
        "Lignes de la commande/Description4": f"Date de fin de contrat : {date_fin}" if date_fin else "Date de fin de contrat : ",
        "Lignes de la commande/Description5": f"Durée du contrat (en mois) : {duree}" if duree else "Durée du contrat (en mois) : ",
        "Lignes de la commande/Description6": f"Consommation annuelle de référence (CAR) : {car}  MWh/an" if car else "Consommation annuelle de référence (CAR) : 0  MWh/an",
        "Lignes de la commande/Quantité": "",
        "Lignes de la commande/Prix Unitaire": "",
        "Lignes de la commande/Distribution Analytique": "",
    }


def dedupe_client_on_product_rows(df_final):
    last_client = None
    for idx in df_final.index:
        display_type = to_str(df_final.at[idx, "Lignes de la commande/Type d'affichage"])
        client = to_str(df_final.at[idx, "Client"])
        if display_type == "NOTE":
            df_final.at[idx, "Client"] = ""
            continue
        if client == last_client:
            df_final.at[idx, "Client"] = ""
        elif client:
            last_client = client
    return df_final


# =========================
# TRANSFORMATION PRINCIPALE
# =========================
def transform_import_odoo(input_excel_path, referentiel_path, output_excel_path=None, sheet_name="Copie de Import Odoo"):
    input_excel_path = Path(input_excel_path)
    referentiel_path = Path(referentiel_path)

    if output_excel_path is None:
        output_excel_path = input_excel_path.with_name(f"{input_excel_path.stem}_transforme.xlsx")
    else:
        output_excel_path = Path(output_excel_path)

    # Lecture brute
    df = pd.read_excel(input_excel_path, sheet_name=sheet_name, header=None, dtype=object)

    # Reprise de la logique historique Power Query
    df = df.iloc[15:].copy().reset_index(drop=True)
    df.columns = [f"Column{i}" for i in range(1, len(df.columns) + 1)]

    if "Column17" in df.columns:
        df = df.rename(columns={"Column17": "Type"})

    headers = df.iloc[0].tolist()
    df = df.iloc[1:].copy().reset_index(drop=True)
    df.columns = headers

    for col in ["Column1", "Column4"]:
        if col in df.columns:
            df = df.drop(columns=[col])

    rename_map = {
        "Fournisseur ↑": "Client",
        "Date de Début Souhaitée": "Lignes de la commande/Description3",
        "Date de fin contrat CE": "Lignes de la commande/Description4",
        "Account Name": "Lignes de la commande/Description1",
        "Compteur": "PDL",
        "Durée": "Lignes de la commande/Description5",
        "CAR validée fournisseur (MWh)": "Lignes de la commande/Description6",
        "Propriétaire de l'opportunité": "Vendeur",
        "Energie": "Lignes de la commande/Produit",
        "Pourcentage rétrocession": "Pourcentage.1",
        "Gestionnaire": "Apporteur d'affaire",
        "Prévisionnel commision": "Lignes de la commande/Prix Unitaire",
        "Date de signature": "Lignes de la commande/Date de signature",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # suppression des lignes totalement vides sur le périmètre utile
    useful_cols = [
        c for c in [
            "Client",
            "Lignes de la commande/Produit",
            "Lignes de la commande/Description1",
            "PDL",
            "Lignes de la commande/Prix Unitaire",
            "Pourcentage.1",
            "Apporteur d'affaire",
            "Vendeur",
            "Type",
        ] if c in df.columns
    ]
    if useful_cols:
        mask_not_empty = df[useful_cols].fillna("").astype(str).apply(lambda s: s.str.strip()).ne("").any(axis=1)
        df = df.loc[mask_not_empty].copy()

    # suppression des 4 dernières lignes parasites comme le script historique
    if len(df) >= 4:
        df = df.iloc[:-4].copy()

    # normalisation des champs numériques sensibles
    if "Lignes de la commande/Prix Unitaire" in df.columns:
        df["Lignes de la commande/Prix Unitaire"] = pd.to_numeric(
            df["Lignes de la commande/Prix Unitaire"].astype(str).str.replace(",", ".", regex=False),
            errors="coerce"
        ).fillna(0)

    if "Pourcentage.1" not in df.columns:
        df["Pourcentage.1"] = 0
    df["Pourcentage.1"] = df["Pourcentage.1"].apply(lambda x: parse_percent_to_int(x, default=0))

    maps, _ = load_referentiel(referentiel_path)
    df = build_codes_from_referentiel(df, maps)

    # construction des lignes au format exact du fichier macro:
    # 1 ligne produit puis 1 ligne note
    output_rows = []
    for _, row in df.iterrows():
        output_rows.append(build_product_row(row))
        output_rows.append(build_note_row(row))

    final_columns = [
        "Client",
        "Lignes de la commande/Produit",
        "Lignes de la commande/Type d'affichage",
        "Lignes de la commande/Description1.1",
        "Lignes de la commande/Description2",
        "Lignes de la commande/Description",
        "Lignes de la commande/Description3",
        "Lignes de la commande/Description4",
        "Lignes de la commande/Description5",
        "Lignes de la commande/Description6",
        "Lignes de la commande/Quantité",
        "Lignes de la commande/Prix Unitaire",
        "Lignes de la commande/Distribution Analytique",
    ]

    df_final = pd.DataFrame(output_rows, columns=final_columns)
    df_final = dedupe_client_on_product_rows(df_final)

    df_final.to_excel(output_excel_path, index=False)
    return df_final, output_excel_path


if __name__ == "__main__":
    fichier_source = "Salesforce.xlsx"
    fichier_referentiel = "Comptes_analytiques.xlsx"
    fichier_sortie = "Salesforce_transforme.xlsx"

    df_resultat, chemin_sortie = transform_import_odoo(
        input_excel_path=fichier_source,
        referentiel_path=fichier_referentiel,
        output_excel_path=fichier_sortie,
        sheet_name="Copie de Import Odoo",
    )

    print(f"Transformation terminée : {chemin_sortie}")
    print(f"Nombre de lignes générées : {len(df_resultat)}")
