#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
from datetime import datetime
from typing import List

import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# RÉPERTOIRES & PARAMÈTRES
# ============================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.environ.get("INPUT_FOLDER", SCRIPT_DIR)
OUTPUT_DIR = os.environ.get("OUTPUT_FOLDER", os.path.join(SCRIPT_DIR, "merged_files"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Année du fichier final
FINAL_YEAR = os.environ.get("FINAL_YEAR") or str(datetime.now().year)

# Colonnes à supprimer
DROP_COLUMNS = {"Numéro", "Nom", "Importante", "source_fichier"}

# Largeurs colonnes pour mise en forme
COLUMN_WIDTHS = {
    "Question": 140,
    "Type de question": 24,
    "Réponse": 70,
    "Valide": 10,
    "Feedback": 140,
}

# ============================================================
# UTILITAIRES
# ============================================================
def read_csv_robust(path: str) -> pd.DataFrame:
    """Lecture robuste CSV avec détection séparateur et encodage."""
    last_err = None
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return pd.read_csv(path, sep=None, engine="python", encoding=enc, dtype=str)
        except Exception as e:
            last_err = e
    raise ValueError(f'Lecture CSV "{os.path.basename(path)}" impossible : {last_err}')

def build_final_name_from_first_csv(csv_filename: str) -> str:
    """
    Exemple : quiz-3001.csv → 3001_Antirétroviraux (1)_biblio_2025.xlsx
    """
    base_noext = os.path.splitext(csv_filename)[0]
    parts = base_noext.split("_", 1)
    if len(parts) == 2:
        numero, titre = parts
    else:
        numero, titre = parts[0], "quiz"
    titre_slug = titre.strip()
    m = re.match(r"^(\d{4})", numero)
    numero_clean = m.group(1) if m else numero.strip()
    return f"{numero_clean}_{titre_slug}_biblio_{FINAL_YEAR}.xlsx"

def format_sheet(ws):
    """Mise en forme : freeze, filtre, largeurs, wrap text, bordures."""
    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(ws.max_column)
    ws.auto_filter.ref = f"A1:{last_col_letter}1"

    headers = [c.value if c.value is not None else "" for c in ws[1]]
    header_to_idx = {str(name): i + 1 for i, name in enumerate(headers)}

    for col_name, width in COLUMN_WIDTHS.items():
        idx = header_to_idx.get(col_name)
        if idx:
            ws.column_dimensions[get_column_letter(idx)].width = width

    for col_name in {"Question", "Réponse", "Feedback"}:
        idx = header_to_idx.get(col_name)
        if idx:
            for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

    thin_border = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin_border

# ============================================================
# PIPELINE PRINCIPAL
# ============================================================
def main():
    csv_files: List[str] = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".csv")]
    if not csv_files:
        print(f"Aucun CSV trouvé dans {INPUT_DIR}")
        return

    csv_files_sorted = sorted(csv_files)
    first_csv = csv_files_sorted[0]
    final_name = build_final_name_from_first_csv(first_csv)
    final_path = os.path.join(OUTPUT_DIR, final_name)

    # Fusion CSV
    frames: List[pd.DataFrame] = []
    for name in csv_files_sorted:
        df = read_csv_robust(os.path.join(INPUT_DIR, name))
        frames.append(df)

    fusion = pd.concat(frames, axis=0, ignore_index=True, sort=False)

    # Supprimer colonnes inutiles
    fusion = fusion[[c for c in fusion.columns if c not in DROP_COLUMNS]]

    # Supprimer doublons
    fusion.drop_duplicates(inplace=True)

    # Créer classeur avec 2 feuilles
    wb = Workbook()
    ws_quiz = wb.active
    ws_quiz.title = "QUIZ"

    # Écrire données QUIZ
    ws_quiz.append(list(fusion.columns))
    for row in fusion.itertuples(index=False, name=None):
        ws_quiz.append(row)
    format_sheet(ws_quiz)

    # Ajouter feuille vide BIBLIOGRAPHIE
    ws_biblio = wb.create_sheet(title="BIBLIOGRAPHIE")
    # Optionnel : ajouter en-tête
    # ws_biblio.append(["Référence"])

    # Sauvegarde
    wb.save(final_path)
    wb.close()

    # Export nom final pour transformationxlsx.py
    os.environ["FINAL_XLSX_NAME"] = final_name
    print(f"✅ Fichier final généré : {final_path}")

if __name__ == "__main__":
    main()