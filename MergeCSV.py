#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
from datetime import datetime
from typing import List

import pandas as pd
from openpyxl import Workbook
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
FINAL_YEAR = os.environ.get("FINAL_YEAR") or "2025"  # fixe 2025 par défaut, change si besoin

# Colonnes à supprimer de la feuille QUIZ
DROP_COLUMNS = {"Numéro", "Nom", "Importante", "source_fichier"}

# Largeurs colonnes pour mise en forme de QUIZ
COLUMN_WIDTHS = {
    "Question": 140,
    "Type de question": 24,
    "Réponse": 70,
    "Valide": 10,
    "Feedback": 140,
    "NbCar Feedback": 16,
}

# ============================================================
# UTILITAIRES
# ============================================================
def read_csv_robust(path: str) -> pd.DataFrame:
    """Lecture robuste CSV (séparateur/encodage auto)."""
    last_err = None
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            return pd.read_csv(path, sep=None, engine="python", encoding=enc, dtype=str)
        except Exception as e:
            last_err = e
    raise ValueError(f'Lecture CSV "{os.path.basename(path)}" impossible : {last_err}')

def sanitize_title(title: str) -> str:
    """
    Autorise lettres (avec accents), chiffres, espaces, (), _ et -.
    Remplace les autres caractères par un espace.
    """
    # Remplace séparateurs multiples par un espace simple
    title = re.sub(r"[^\w\s\(\)\-àâäéèêëîïôöùûüçÀÂÄÉÈÊËÎÏÔÖÙÛÜÇ]", " ", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title

def build_final_name_from_csv_content(df_first: pd.DataFrame) -> str:
    """
    Construit '3001_Antirétroviraux (1)_biblio_2025.xlsx'
    depuis les colonnes 'Numéro' et 'Nom' du 1er CSV.
    """
    numero = None
    titre = None
    # essaie de lire la première ligne non vide
    if "Numéro" in df_first.columns:
        numero = (df_first["Numéro"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().head(1).tolist() or [None])[0]
    if "Nom" in df_first.columns:
        titre = (df_first["Nom"].dropna().astype(str).str.strip().replace("", pd.NA).dropna().head(1).tolist() or [None])[0]

    # fallback si champs manquants
    if not numero:
        numero = "0000"
    if not titre:
        titre = "quiz"

    titre = sanitize_title(titre)
    # sécurise le numéro (garde les 4 premiers chiffres si présents)
    m = re.match(r"^(\d{4})", str(numero))
    numero_clean = m.group(1) if m else str(numero).strip()

    return f"{numero_clean}_{titre}_biblio_{FINAL_YEAR}.xlsx"

def format_sheet(ws):
    """Mise en forme QUIZ : freeze, filtre, largeurs, wrap text, bordures fines."""
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

    # Lit le 1er CSV pour le nommage
    df_first = read_csv_robust(os.path.join(INPUT_DIR, csv_files_sorted[0]))
    final_name = build_final_name_from_csv_content(df_first)
    final_path = os.path.join(OUTPUT_DIR, final_name)

    # Nettoie les anciens xlsx pour repartir propre
    for name in os.listdir(OUTPUT_DIR):
        p = os.path.join(OUTPUT_DIR, name)
        if os.path.isfile(p) and name.lower().endswith(".xlsx"):
            try:
                os.remove(p)
            except Exception:
                pass

    # Fusionne tous les CSV
    frames: List[pd.DataFrame] = [df_first]
    for name in csv_files_sorted[1:]:
        frames.append(read_csv_robust(os.path.join(INPUT_DIR, name)))
    fusion = pd.concat(frames, axis=0, ignore_index=True, sort=False)

    # Supprime colonnes inutiles
    keep_cols = [c for c in fusion.columns if c not in DROP_COLUMNS]
    fusion = fusion[keep_cols]

    # Supprime doublons sur les colonnes restantes
    fusion.drop_duplicates(inplace=True)

    # Ajoute une colonne "NbCar Feedback" si "Feedback" existe
    if "Feedback" in fusion.columns:
        fusion["NbCar Feedback"] = fusion["Feedback"].fillna("").astype(str).map(len)

    # Ordre des colonnes : celles-ci d'abord si présentes
    preferred = ["Question", "Type de question", "Réponse", "Valide", "Feedback", "NbCar Feedback"]
    ordered = [c for c in preferred if c in fusion.columns] + [c for c in fusion.columns if c not in preferred]
    fusion = fusion[ordered]

    # Crée le classeur final avec 2 feuilles
    wb = Workbook()
    ws_quiz = wb.active
    ws_quiz.title = "QUIZ"

    # Écrit les données QUIZ
    ws_quiz.append(list(fusion.columns))
    for row in fusion.itertuples(index=False, name=None):
        ws_quiz.append(row)
    format_sheet(ws_quiz)

    # Ajoute la feuille BIBLIOGRAPHIE vide
    wb.create_sheet(title="BIBLIOGRAPHIE")

    # Sauvegarde
    wb.save(final_path)
    wb.close()

    # Expose le nom pour transformationxlsx.py (si utilisé)
    os.environ["FINAL_XLSX_NAME"] = final_name
    print(f"✅ Fichier final généré : {final_path}")

if __name__ == "__main__":
    main()