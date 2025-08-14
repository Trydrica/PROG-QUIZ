#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import unicodedata
from typing import List

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter

# =========================
# Chemins / paramètres
# =========================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.environ.get("INPUT_FOLDER", SCRIPT_DIR)
OUTPUT_DIR = os.environ.get("OUTPUT_FOLDER", os.path.join(SCRIPT_DIR, "merged_files"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

FINAL_YEAR = os.environ.get("FINAL_YEAR") or "2025"  # mets str(datetime.now().year) si tu veux dynamique

# Colonnes retirées de la feuille QUIZ
DROP_COLUMNS = {"Numéro", "Nom", "Importante", "source_fichier"}

# Largeurs de colonnes EXACTES demandées
FIXED_COLUMN_WIDTHS = {
    "Question": 50,
    "Réponse": 30,
    "Valide": 10,
    "Feedback": 140,
    "NbCar Feedback": 16,
}

# Colonnes à fusionner visuellement si identiques consécutives
MERGE_VISUAL_COLS = ["Question", "Feedback"]

# Hauteur des lignes en points (on applique 150 comme demandé)
ROW_HEIGHT_POINTS = 150.0


# =========================
# Utilitaires
# =========================
def read_csv_robust(path: str) -> pd.DataFrame:
    last_err = None
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return pd.read_csv(path, sep=None, engine="python", encoding=enc, dtype=str)
        except Exception as e:
            last_err = e
    raise ValueError(f'Lecture CSV "{os.path.basename(path)}" impossible : {last_err}')

def sanitize_filename(name: str) -> str:
    # Conserver accents et parenthèses, retirer caractères interdits par le FS
    return re.sub(r'[\\/:*?"<>|]', " ", name).strip()

def build_final_name_from_content(df_first: pd.DataFrame) -> str:
    """Construit '3001_NomDuQuiz_biblio_2025.xlsx' depuis 'Numéro' + 'Nom' du 1er CSV (contenu)."""
    # Numéro (4 premiers chiffres si possible)
    numero = None
    if "Numéro" in df_first.columns:
        s = df_first["Numéro"].dropna().astype(str).str.strip().replace("", pd.NA).dropna()
        if not s.empty:
            numero = s.iloc[0]
    if not numero:
        numero = "0000"
    m = re.match(r"^(\d{4})", str(numero))
    numero_clean = m.group(1) if m else str(numero).strip()

    # Titre — garder accents/parenthèses, compresser espaces
    titre = None
    if "Nom" in df_first.columns:
        s = df_first["Nom"].dropna().astype(str).str.strip().replace("", pd.NA).dropna()
        if not s.empty:
            titre = s.iloc[0]
    if not titre:
        titre = "quiz"
    titre = re.sub(r"\s+", " ", str(titre)).strip()

    final_name = f"{numero_clean}_{titre}_biblio_{FINAL_YEAR}.xlsx"
    return sanitize_filename(final_name)

def _normalize_text(s: str) -> str:
    if s is None:
        return ""
    s = unicodedata.normalize("NFKC", str(s))
    s = s.replace("’", "'").replace("‘", "'").replace("“", '"').replace("”", '"')
    s = s.replace("–", "-").replace("—", "-")
    s = s.replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s)
    return s.strip()

def normalize_df_key(df: pd.DataFrame) -> pd.Series:
    """Clé de déduplication normalisée sur toutes les colonnes (insensible à la casse)."""
    cols = []
    for c in df.columns:
        cols.append(df[c].fillna("").astype(str).map(_normalize_text).str.lower())
    key = cols[0]
    for i in range(1, len(cols)):
        key = key + "||" + cols[i]
    return key


# =========================
# Écriture / Mise en forme
# =========================
def write_quiz_sheet(wb: Workbook, df: pd.DataFrame) -> None:
    """Crée la feuille QUIZ : données + format + fusion + hauteurs/largeurs + formule LEN."""
    ws = wb.active
    ws.title = "QUIZ"

    # 1) Entête + données
    headers = list(df.columns)
    ws.append(headers)
    for row in df.itertuples(index=False, name=None):
        ws.append(row)

    header_map = {name: i + 1 for i, name in enumerate(headers)}

    # 2) Gel des volets + filtre
    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(ws.max_column)
    ws.auto_filter.ref = f"A1:{last_col_letter}1"

    # 3) Hauteur lignes (2..N) + alignement/wrap
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        ws.row_dimensions[row[0].row].height = ROW_HEIGHT_POINTS
        for cell in row:
            if cell.value is not None:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # 4) Largeurs fixes
    for col_name, width in FIXED_COLUMN_WIDTHS.items():
        idx = header_map.get(col_name)
        if idx:
            ws.column_dimensions[get_column_letter(idx)].width = width

    # 5) Bordures fines
    thin = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin

    # 6) Fusion visuelle (identiques consécutives, normalisées, non vides)
    def merge_consecutive(col_name: str):
        idx = header_map.get(col_name)
        if not idx or ws.max_row < 3:
            return

        def _norm(x):
            if x is None: return ""
            x = unicodedata.normalize("NFKC", str(x))
            x = x.replace("\u00A0", " ")
            x = re.sub(r"\s+", " ", x).strip().lower()
            x = x.replace("’","'").replace("“",'"').replace("”",'"').replace("–","-").replace("—","-")
            return x

        start = None
        prev = None
        for r in range(2, ws.max_row + 1):
            raw = ws.cell(row=r, column=idx).value
            cur = _norm(raw)
            if not cur:
                if start is not None and r - 1 > start:
                    ws.merge_cells(start_row=start, end_row=r - 1, start_column=idx, end_column=idx)
                    ws.cell(row=start, column=idx).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                start, prev = None, None
                continue

            if prev is None:
                start, prev = r, cur
                continue

            if cur == prev:
                continue

            if r - 1 > start:
                ws.merge_cells(start_row=start, end_row=r - 1, start_column=idx, end_column=idx)
                ws.cell(row=start, column=idx).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

            start, prev = r, cur

        if start is not None and ws.max_row > start:
            ws.merge_cells(start_row=start, end_row=ws.max_row, start_column=idx, end_column=idx)
            ws.cell(row=start, column=idx).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    for col in MERGE_VISUAL_COLS:
        merge_consecutive(col)

    # 7) Formule LEN(Feedback) -> "NbCar Feedback"
    headers_now = [c.value if c.value is not None else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
    h2idx = {name: i + 1 for i, name in enumerate(headers_now)}
    c_feedback = h2idx.get("Feedback")
    c_len = h2idx.get("NbCar Feedback")
    if c_feedback and c_len:
        colF = get_column_letter(c_feedback)
        for r in range(2, ws.max_row + 1):
            ws.cell(row=r, column=c_len).value = f"=LEN({colF}{r})"  # Excel FR => NBCAR


# =========================
# Pipeline principal
# =========================
def main():
    csv_files: List[str] = sorted([f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".csv")])
    if not csv_files:
        print(f"Aucun CSV trouvé dans {INPUT_DIR}")
        return

    # Nom final d'après CONTENU du 1er CSV
    df_first = read_csv_robust(os.path.join(INPUT_DIR, csv_files[0]))
    final_name = build_final_name_from_content(df_first)
    final_path = os.path.join(OUTPUT_DIR, final_name)

    # Nettoie anciens .xlsx
    for name in os.listdir(OUTPUT_DIR):
        p = os.path.join(OUTPUT_DIR, name)
        if os.path.isfile(p) and name.lower().endswith(".xlsx"):
            try:
                os.remove(p)
            except Exception:
                pass

    # Fusionne tous les CSV fournis (tu les envoies par paires)
    frames: List[pd.DataFrame] = [df_first]
    for name in csv_files[1:]:
        frames.append(read_csv_robust(os.path.join(INPUT_DIR, name)))
    fusion = pd.concat(frames, axis=0, ignore_index=True, sort=False)

    # Retire colonnes inutiles
    fusion = fusion[[c for c in fusion.columns if c not in DROP_COLUMNS]]

    # Dédup “intelligente”
    key = normalize_df_key(fusion)
    fusion = fusion.loc[~key.duplicated(keep="first")].reset_index(drop=True)

    # Ajout NbCar Feedback si Feedback existe
    if "Feedback" in fusion.columns and "NbCar Feedback" not in fusion.columns:
        fusion["NbCar Feedback"] = 0

    # Ordre colonnes
    preferred = ["Question", "Type de question", "Réponse", "Valide", "Feedback", "NbCar Feedback"]
    fusion = fusion[[c for c in preferred if c in fusion.columns] + [c for c in fusion.columns if c not in preferred]]

    # Écriture finale : QUIZ + BIBLIOGRAPHIE vide
    wb = Workbook()
    write_quiz_sheet(wb, fusion)
    wb.create_sheet(title="BIBLIOGRAPHIE")
    wb.save(final_path)
    wb.close()

    os.environ["FINAL_XLSX_NAME"] = final_name
    print(f"✅ Fichier final généré :", final_path)

if __name__ == "__main__":
    main()