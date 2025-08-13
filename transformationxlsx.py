#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
from typing import Dict

from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter

# ============================================================
# RÉPERTOIRE & CIBLE
# ============================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.environ.get("OUTPUT_FOLDER", os.path.join(SCRIPT_DIR, "merged_files"))

# On récupère le nom décidé par MergeCSV.py si dispo,
# sinon on prend le seul .xlsx présent dans OUTPUT_DIR.
FINAL_NAME = os.environ.get("FINAL_XLSX_NAME")
if not FINAL_NAME:
    candidates = [f for f in os.listdir(OUTPUT_DIR) if f.lower().endswith(".xlsx")]
    if not candidates:
        raise FileNotFoundError(f"Aucun .xlsx trouvé dans {OUTPUT_DIR}")
    # S'il y en a plusieurs (peu probable car on a nettoyé), on prend le premier trié
    FINAL_NAME = sorted(candidates)[0]

FINAL_PATH = os.path.join(OUTPUT_DIR, FINAL_NAME)
if not os.path.exists(FINAL_PATH):
    raise FileNotFoundError(f"Le fichier final attendu est introuvable : {FINAL_PATH}")

# ============================================================
# PARAMÈTRES DE PRÉSENTATION
# ============================================================
COLUMN_WIDTHS: Dict[str, float] = {
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

BORDER_THIN = Border(
    left=Side(style="thin", color="CCCCCC"),
    right=Side(style="thin", color="CCCCCC"),
    top=Side(style="thin", color="CCCCCC"),
    bottom=Side(style="thin", color="CCCCCC"),
)

HEADER_FONT = Font(bold=True)

def format_sheet(ws) -> None:
    # Figer l’entête & filtre
    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(ws.max_column)
    ws.auto_filter.ref = f"A1:{last_col_letter}1"

    # En-têtes pour mapping
    headers = [c.value if c.value is not None else "" for c in ws[1]]
    header_to_idx = {str(name): i + 1 for i, name in enumerate(headers)}

    # Largeurs
    for col_name, width in COLUMN_WIDTHS.items():
        idx = header_to_idx.get(col_name)
        if idx:
            ws.column_dimensions[get_column_letter(idx)].width = width

    # Wrap + align top
    for col_name in {"Question", "Réponse", "Feedback"}:
        idx = header_to_idx.get(col_name)
        if idx:
            for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

    # Bordures fines
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = BORDER_THIN

def main() -> None:
    wb = load_workbook(FINAL_PATH)
    for ws in wb.worksheets:
        format_sheet(ws)
    # ÉCRASE LE MÊME FICHIER (pas de *_updated)
    wb.save(FINAL_PATH)
    wb.close()
    print(f"✅ Mise en forme appliquée (écrasement) : {FINAL_PATH}")

if __name__ == "__main__":
    main()