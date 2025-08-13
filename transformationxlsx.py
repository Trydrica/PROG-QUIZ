#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ============================================================
# RÉPERTOIRE & CIBLE
# ============================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.environ.get("OUTPUT_FOLDER", os.path.join(SCRIPT_DIR, "merged_files"))

FINAL_NAME = os.environ.get("FINAL_XLSX_NAME")
if not FINAL_NAME:
    # Si non passé par MergeCSV.py, on prend le seul .xlsx dans OUTPUT_DIR
    candidates = [f for f in os.listdir(OUTPUT_DIR) if f.lower().endswith(".xlsx")]
    if not candidates:
        raise FileNotFoundError(f"Aucun .xlsx trouvé dans {OUTPUT_DIR}")
    FINAL_NAME = sorted(candidates)[0]

FINAL_PATH = os.path.join(OUTPUT_DIR, FINAL_NAME)
if not os.path.exists(FINAL_PATH):
    raise FileNotFoundError(f"Le fichier final attendu est introuvable : {FINAL_PATH}")

# ============================================================
# PARAMÈTRES DE PRÉSENTATION
# ============================================================
COLUMN_WIDTHS = {
    "Question": 140,
    "Type de question": 24,
    "Réponse": 70,
    "Valide": 10,
    "Feedback": 140,
}

def format_sheet(ws):
    """Mise en forme : freeze, filtre, largeurs, wrap text, bordures fines."""
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
# MAIN
# ============================================================
def main():
    wb = load_workbook(FINAL_PATH)
    if "QUIZ" in wb.sheetnames:
        format_sheet(wb["QUIZ"])
    wb.save(FINAL_PATH)
    wb.close()
    print(f"✅ Mise en forme appliquée sur : {FINAL_PATH}")

if __name__ == "__main__":
    main()