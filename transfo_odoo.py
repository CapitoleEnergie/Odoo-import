
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
        return round(float(value), decimals)
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


def to_percent_int(value):
    if pd.isna(value) or value == "":
        return None
    try:
        # accepte 0.75 ou 75
        num = float(value)
        if num <= 1:
            num = num * 100
        return int(round(num))
    except Exception:
        return None


def clean_id(value):
    txt = to_str(value)
    if not txt:
        return ""
    try:
        return str(int(float(txt)))
    except Exception:
        return txt


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
    ref["ID"] = ref["ID"].apply(clean_id)

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
    bu_interne = get_ref_id(maps, "01 - BU", "02 - Courtage Interne", default="48")

    df["AA_interne"] = (
        df["Apporteur d'affaire"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
        .eq("CAPITOLE ENERGIE")
    )
    df["Gestionnaire-BU"] = np.where(df["AA_interne"], str(bu_interne), str(bu_externe))

    # Type
    df["Type"] = df["Type"].apply(
        lambda x: get_ref_id(maps, "02 - Niveau", x, default=None, aliases=ALIASES_02)
    )

    # Produit
    df["Valeur:Produit"] = df["Lignes de la commande/Produit"].apply(
        lambda x: get_ref_id(maps, "03 - Niveau", x, default=None, aliases=ALIASES_03)
    )

    # Codes apporteur / vendeur
    def map_codes_ok(row):
        apporteur = row.get("Apporteur d'affaire", "")
        vendeur = row.get("Vendeur", "")

        code_apporteur = get_ref_id(
            maps, "04 - Niveau", apporteur, default=None, aliases=ALIASES_04
        )
        if code_apporteur is not None:
            return clean_id(code_apporteur)

        code_vendeur = get_ref_id(
            maps, "04 - Niveau", vendeur, default=None, aliases=ALIASES_04
        )
        if code_vendeur is not None:
            return clean_id(code_vendeur)

        return ""

    df["Codes OK"] = df.apply(map_codes_ok, axis=1)

    ce_aa = get_ref_id(
        maps,
        "04 - Niveau",
        "Capitole Energie AA",
        default="76",
        aliases=ALIASES_04,
    )
    df["Capitole Energie AA"] = np.where(~df["AA_interne"], clean_id(ce_aa), "")
    return df


def build_distribution_key(*ids):
    ids = [clean_id(v) for v in ids if clean_id(v)]
    if len(ids) < 4:
        return ""
    return ",".join(ids)


def build_analytics(row):
    bu = row.get("Gestionnaire-BU", "")
    type_ = row.get("Type", "")
    produit = row.get("Valeur:Produit", "")
    code = row.get("Codes OK", "")
    ce = row.get("Capitole Energie AA", "")
    pct = row.get("Pourcentage", None)

    if not bu or not type_ or not produit or not code:
        return "", "Code analytique incomplet"
    if pct is None:
        return "", "Pourcentage rétrocession manquant ou invalide"

    # borne de sécurité
    pct = max(0, min(100, int(pct)))

    distribution = {}
    key_main = build_distribution_key(bu, type_, produit, code)
    if not key_main:
        return "", "Clé analytique principale invalide"
    distribution[key_main] = pct

    # Pour AA externes, on complète avec Capitole Energie AA
    if ce:
        pct_diff = 100 - pct
        if pct_diff > 0:
            key_ce = build_distribution_key(bu, type_, produit, ce)
            if not key_ce:
                return "", "Clé analytique Capitole Energie AA invalide"
            distribution[key_ce] = pct_diff

    # json.dumps garantit un JSON valide
    return json.dumps(distribution, ensure_ascii=False, separators=(",", ":")), ""


# =========================
# TRANSFORMATION PRINCIPALE
# =========================
def transform_import_odoo(
    input_excel_path,
    referentiel_path,
    output_excel_path=None,
    sheet_name="Copie de Import Odoo",
):
    input_excel_path = Path(input_excel_path)
    referentiel_path = Path(referentiel_path)

    if output_excel_path is None:
        output_excel_path = input_excel_path.with_name(f"{input_excel_path.stem}_transforme.xlsx")
    else:
        output_excel_path = Path(output_excel_path)

    df = pd.read_excel(input_excel_path, sheet_name=sheet_name, header=None, dtype=object)

    # Reproduction du comportement Power Query : on garde à partir de la ligne 16
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
        "Compteur": "Lignes de la commande/Description2",
        "Durée": "Lignes de la commande/Description5",
        "CAR validée fournisseur (MWh)": "Lignes de la commande/Description6",
        "Propriétaire de l'opportunité": "Vendeur",
        "Energie": "Lignes de la commande/Produit",
        "Pourcentage rétrocession": "Pourcentage.1",
        "Gestionnaire": "Apporteur d'affaire",
        "Prévisionnel commision": "Lignes de la commande/Prix Unitaire",
        "Date de signature": "Lignes de la commande/Date de signature",
        "Column19": "Lignes de la commande/Distribution Analytique",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    maps, ref_df = load_referentiel(referentiel_path)

    if "Lignes de la commande/Description2" in df.columns:
        df = df.rename(columns={"Lignes de la commande/Description2": "PDL"})
        df["Lignes de la commande/Description2"] = (
            "PDL : " + df["PDL"].fillna("").astype(str).str.strip()
        )
        df = df.drop(columns=["PDL"])

    if "Lignes de la commande/Date de signature" in df.columns:
        df["Lignes de la commande/Description"] = (
            "Date de signature : "
            + df["Lignes de la commande/Date de signature"].fillna("").astype(str).str.strip()
        )

    if "Lignes de la commande/Description1" in df.columns:
        df["Lignes de la commande/Description1"] = (
            "RAISON SOCIALE : "
            + df["Lignes de la commande/Description1"].fillna("").astype(str).str.strip()
        )

    if "Lignes de la commande/Description5" in df.columns:
        df["Lignes de la commande/Description5"] = (
            "Durée du contrat (en mois) : "
            + df["Lignes de la commande/Description5"].fillna("").astype(str).str.strip()
        )

    if "Lignes de la commande/Description6" in df.columns:
        df["Lignes de la commande/Description6_num"] = (
            df["Lignes de la commande/Description6"].apply(safe_round)
        )
        df["Lignes de la commande/Description6"] = (
            "Consommation annuelle de référence (CAR) : "
            + df["Lignes de la commande/Description6_num"].apply(lambda x: format_number_fr(x, 3))
            + " MWh/an"
        )

    if "Lignes de la commande/Description3" in df.columns:
        df["Lignes de la commande/Description3"] = (
            "Date de début contrat : "
            + df["Lignes de la commande/Description3"].fillna("").astype(str).str.strip()
        )

    if "Lignes de la commande/Description4" in df.columns:
        df["Lignes de la commande/Description4"] = (
            "Date de fin de contrat : "
            + df["Lignes de la commande/Description4"].fillna("").astype(str).str.strip()
        )

    df = build_codes_from_referentiel(df, maps)

    # Suppression des 4 dernières lignes comme dans Power Query
    if len(df) >= 4:
        df = df.iloc[:-4].copy()

    df["Pourcentage.1"] = pd.to_numeric(df["Pourcentage.1"], errors="coerce")
    df["Pourcentage"] = df["Pourcentage.1"].apply(to_percent_int)

    analytics = df.apply(build_analytics, axis=1, result_type="expand")
    analytics.columns = ["Lignes de la commande/Distribution Analytique", "Erreur analytique"]
    df = pd.concat([df, analytics], axis=1)

    # On ne garde que les lignes valides pour éviter un import Odoo cassé
    df_valid = df[df["Erreur analytique"].eq("")].copy()
    df_errors = df[df["Erreur analytique"].ne("")].copy()

    # Lignes produit facturables : surtout pas de Type d'affichage = NOTE
    df_valid["Lignes de la commande/Quantité"] = "1"
    df_valid["Lignes de la commande/Description1.1"] = (
        "Contrat Energie " + df_valid["Lignes de la commande/Produit"].fillna("").astype(str).str.strip()
    )

    cols_to_drop = [
        "Valeur:Produit",
        "Gestionnaire-BU",
        "Codes OK",
        "Pourcentage",
        "Capitole Energie AA",
        "AA_interne",
        "Lignes de la commande/Distribution Analytique1",
        "Analytics fusion",
        "Analytics OK pour fusion",
        "2me Ligne pour fusion",
        "2me Ligne",
        "Vendeur",
        "Type",
        "Apporteur d'affaire",
        "Pourcentage.1",
        "Montant rétrocession Total",
        "Mois pour facturation ↑",
        "Lignes de la commande/Date de signature",
        "Lignes de la commande/Description6_num",
    ]
    cols_to_drop = [c for c in cols_to_drop if c in df_valid.columns]
    df_valid = df_valid.drop(columns=cols_to_drop)

    final_columns = [
        "Client",
        "Lignes de la commande/Produit",
        "Lignes de la commande/Description1.1",
        "Lignes de la commande/Description1",
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
    final_columns_existing = [c for c in final_columns if c in df_valid.columns]
    df_final = df_valid[final_columns_existing].copy()

    # Export principal
    df_final.to_excel(output_excel_path, index=False)

    # Export rejet si besoin
    if not df_errors.empty:
        error_path = output_excel_path.with_name(f"{output_excel_path.stem}_erreurs.xlsx")
        df_errors.to_excel(error_path, index=False)

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
