#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import csv
from io import BytesIO
from typing import List
from datetime import datetime

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# RÉPERTOIRES & PARAMÈTRES
# ============================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.environ.get("INPUT_FOLDER", SCRIPT_DIR)
OUTPUT_DIR = os.environ.get("OUTPUT_FOLDER", os.path.join(SCRIPT_DIR, "merged_files"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Année de création : FINAL_YEAR (env) ou année courante
FINAL_YEAR = os.environ.get("FINAL_YEAR")
if not FINAL_YEAR:
    FINAL_YEAR = str(datetime.now().year)

# ============================================================
# UTILITAIRES
# ============================================================
def read_csv_robust(path: str) -> pd.DataFrame:
    """
    Lecture robuste d'un CSV :
      - détection auto du séparateur via sep=None + engine='python'
      - essais d'encodage : utf-8-sig, utf-8, latin-1, cp1252
      - dtype=str pour préserver les zéros initiaux
    """
    last_err = None
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return pd.read_csv(path, sep=None, engine="python", encoding=enc, dtype=str)
        except Exception as e:
            last_err = e
    raise ValueError(f'Lecture CSV "{os.path.basename(path)}" impossible : {last_err}')

def build_final_name_from_first_csv(csv_filename: str) -> str:
    """
    csv_filename ex: '1001_IEC et sartans.csv'
    -> '1001_IEC_et_sartans_biblio_2025.xlsx'  (année selon FINAL_YEAR)
    """
    base_noext = os.path.splitext(csv_filename)[0]
    parts = base_noext.split("_", 1)
    if len(parts) == 2:
        numero, titre = parts
    else:
        numero, titre = parts[0], "quiz"
    # Nettoyage du titre -> underscores
    titre_slug = "_".join(t.strip() for t in titre.strip().split())
    # Sécurisation basique du numéro (garde 4 premiers chiffres si présents)
    m = re.match(r"^(\d{4})", numero)
    numero_clean = m.group(1) if m else numero.strip()
    return f"{numero_clean}_{titre_slug}_biblio_{FINAL_YEAR}.xlsx"

def write_xlsx_with_format(df: pd.DataFrame, out_path: str, sheet_name: str = "Fusion") -> None:
    """
    Écrit un DataFrame en .xlsx + quelques mises en forme utiles :
      - Freeze en-tête (A2)
      - Auto-filtre
      - Largeurs de colonnes usuelles si présentes
      - Wrap text pour colonnes longues (Question/Réponse/Feedback)
    """
    # Colonnes préférées en tête si elles existent
    preferred = [
        "Numéro", "Nom", "Question", "Type de question",
        "Réponse", "Valide", "Importante", "Feedback", "source_fichier"
    ]
    cols = [c for c in preferred if c in df.columns] + [c for c in df.columns if c not in preferred]
    df = df[cols]

    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet_name)

    wb = load_workbook(out_path)
    ws = wb[sheet_name]

    # Figer l’entête & filtre
    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(ws.max_column)
    ws.auto_filter.ref = f"A1:{last_col_letter}1"

    # Largeurs courantes (adapter si besoin)
    column_widths = {
        "Numéro": 10,
        "Nom": 30,
        "Question": 140,
        "Type de question": 24,
        "Réponse": 70,
        "Valide": 10,
        "Importante": 14,
        "Feedback": 140,
        "source_fichier": 28,
    }
    headers = [c.value if c.value is not None else "" for c in ws[1]]
    header_to_idx = {str(name): i + 1 for i, name in enumerate(headers)}
    for col_name, width in column_widths.items():
        idx = header_to_idx.get(col_name)
        if idx:
            ws.column_dimensions[get_column_letter(idx)].width = width

    # Wrap text sur colonnes longues
    for col_name in {"Question", "Réponse", "Feedback"}:
        idx = header_to_idx.get(col_name)
        if idx:
            for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(out_path)
    wb.close()

# ============================================================
# TRAITEMENT — UN SEUL FICHIER FINAL
# ============================================================
def main() -> None:
    csv_files: List[str] = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".csv")]
    if not csv_files:
        print(f"Aucun .csv trouvé dans : {INPUT_DIR}")
        return

    csv_files_sorted = sorted(csv_files)
    first_csv = csv_files_sorted[0]
    final_name = build_final_name_from_first_csv(first_csv)
    final_path = os.path.join(OUTPUT_DIR, final_name)

    # Nettoie d’anciens xlsx pour repartir propre
    for name in os.listdir(OUTPUT_DIR):
        p = os.path.join(OUTPUT_DIR, name)
        if os.path.isfile(p) and name.lower().endswith(".xlsx"):
            try:
                os.remove(p)
            except Exception:
                pass

    frames: List[pd.DataFrame] = []
    for name in csv_files_sorted:
        src_path = os.path.join(INPUT_DIR, name)
        df = read_csv_robust(src_path)
        # Conserver la source (utile pour audit)
        df.insert(0, "source_fichier", name)
        frames.append(df)

    fusion = pd.concat(frames, axis=0, ignore_index=True, sort=False)
    write_xlsx_with_format(fusion, final_path, sheet_name="Fusion")

    # Expose le nom choisi pour le script de transformation (optionnel)
    os.environ["FINAL_XLSX_NAME"] = final_name

    print(f"✅ Fichier final généré : {final_path}")

if __name__ == "__main__":
    main()