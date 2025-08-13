#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import unicodedata
from datetime import datetime
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
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "merged_files") if not os.environ.get("OUTPUT_FOLDER") else os.environ["OUTPUT_FOLDER"]
os.makedirs(OUTPUT_DIR, exist_ok=True)

FINAL_YEAR = os.environ.get("FINAL_YEAR") or "2025"  # fixe ; mets str(datetime.now().year) si tu veux dynamique

DROP_COLUMNS = {"Numéro", "Nom", "Importante", "source_fichier"}  # supprimées du rendu
COLUMN_WIDTHS = {  # largeur colonnes
    "Question": 140,
    "Type de question": 24,
    "Réponse": 70,
    "Valide": 10,
    "Feedback": 140,
    "NbCar Feedback": 16,
}
MERGE_VISUAL_COLS = ["Question", "Feedback"]  # colonnes à fusionner visuellement si identiques consécutives


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

def sanitize_title(title: str) -> str:
    title = re.sub(r"[^\w\s\(\)\-àâäéèêëîïôöùûüçÀÂÄÉÈÊËÎÏÔÖÙÛÜÇ]", " ", str(title))
    title = re.sub(r"\s+", " ", title).strip()
    return title

def build_final_name_from_content(df_first: pd.DataFrame) -> str:
    # Numéro
    numero = None
    if "Numéro" in df_first.columns:
        s = df_first["Numéro"].dropna().astype(str).str.strip().replace("", pd.NA).dropna()
        if not s.empty:
            numero = s.iloc[0]
    if not numero:
        numero = "0000"
    m = re.match(r"^(\d{4})", str(numero))
    numero_clean = m.group(1) if m else str(numero).strip()
    # Titre
    titre = None
    if "Nom" in df_first.columns:
        s = df_first["Nom"].dropna().astype(str).str.strip().replace("", pd.NA).dropna()
        if not s.empty:
            titre = s.iloc[0]
    if not titre:
        titre = "quiz"
    titre = sanitize_title(titre)
    return f"{numero_clean}_{titre}_biblio_{FINAL_YEAR}.xlsx"

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
    # clé de dédup sur TOUTES les colonnes restantes (normalisées)
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
    ws = wb.active
    ws.title = "QUIZ"

    # Entête + données
    headers = list(df.columns)
    ws.append(headers)
    for row in df.itertuples(index=False, name=None):
        ws.append(row)

    # Format de base
    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(ws.max_column)
    ws.auto_filter.ref = f"A1:{last_col_letter}1"

    header_to_idx = {name: i + 1 for i, name in enumerate(headers)}
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

    thin = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin

    # Fusion VISUELLE des cellules identiques consécutives (normalisées, non vides)
    def merge_consecutive_by_name(col_name: str):
        idx = header_to_idx.get(col_name)
        if not idx or ws.max_row < 3:
            return
        start = None
        prev_norm = None
        for r in range(2, ws.max_row + 1):
            raw = ws.cell(row=r, column=idx).value
            norm = _normalize_text(raw).lower()
            if not norm:  # on ne fusionne pas les vides
                # clôt une éventuelle séquence
                if start is not None and r - 1 > start:
                    ws.merge_cells(start_row=start, end_row=r - 1, start_column=idx, end_column=idx)
                    ws.cell(row=start, column=idx).alignment = Alignment(wrap_text=True, vertical="top")
                start = None
                prev_norm = None
                continue
            if prev_norm is None:  # démarre
                start = r
                prev_norm = norm
                continue
            if norm == prev_norm:
                # on continue la séquence
                continue
            # valeur différente -> fusionner la séquence précédente si longueur >= 2
            if r - 1 > start:
                ws.merge_cells(start_row=start, end_row=r - 1, start_column=idx, end_column=idx)
                ws.cell(row=start, column=idx).alignment = Alignment(wrap_text=True, vertical="top")
            start = r
            prev_norm = norm
        # fin de colonne : fusionner la dernière séquence
        if start is not None and ws.max_row > start:
            ws.merge_cells(start_row=start, end_row=ws.max_row, start_column=idx, end_column=idx)
            ws.cell(row=start, column=idx).alignment = Alignment(wrap_text=True, vertical="top")

    for col in MERGE_VISUAL_COLS:
        merge_consecutive_by_name(col)

def add_len_formula_on_feedback(wb: Workbook) -> None:
    ws = wb["QUIZ"]
    headers = [c.value if c.value is not None else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
    header_to_idx = {name: i + 1 for i, name in enumerate(headers)}
    c_feedback = header_to_idx.get("Feedback")
    c_len = header_to_idx.get("NbCar Feedback")
    if not c_feedback or not c_len:
        return
    from openpyxl.utils import get_column_letter
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

    # Nom final d'après le CONTENU du 1er CSV
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

    # Fusionner tous les CSV fournis (tu enverras 2 par 2)
    frames: List[pd.DataFrame] = [df_first]
    for name in csv_files[1:]:
        frames.append(read_csv_robust(os.path.join(INPUT_DIR, name)))
    fusion = pd.concat(frames, axis=0, ignore_index=True, sort=False)

    # Supprimer colonnes inutiles
    fusion = fusion[[c for c in fusion.columns if c not in DROP_COLUMNS]]

    # Déduplication intelligente (normalisée sur toutes colonnes restantes)
    key = normalize_df_key(fusion)
    fusion = fusion.loc[~key.duplicated(keep="first")].reset_index(drop=True)

    # Ajout colonne NbCar Feedback (formule sera posée ensuite)
    if "Feedback" in fusion.columns and "NbCar Feedback" not in fusion.columns:
        fusion["NbCar Feedback"] = 0

    # Ordre des colonnes
    preferred = ["Question", "Type de question", "Réponse", "Valide", "Feedback", "NbCar Feedback"]
    fusion = fusion[[c for c in preferred if c in fusion.columns] + [c for c in fusion.columns if c not in preferred]]

    # Écriture + format + fusion visuelle + formules
    wb = Workbook()
    write_quiz_sheet(wb, fusion)
    wb.create_sheet(title="BIBLIOGRAPHIE")
    add_len_formula_on_feedback(wb)
    wb.save(final_path)
    wb.close()

    os.environ["FINAL_XLSX_NAME"] = final_name
    print(f"✅ Fichier final généré : {final_path}")

if __name__ == "__main__":
    main()