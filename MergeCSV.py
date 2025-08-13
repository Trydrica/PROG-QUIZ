#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import unicodedata
from datetime import datetime
from typing import List, Optional

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ============================================================
# RÉPERTOIRES & PARAMÈTRES
# ============================================================
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.environ.get("INPUT_FOLDER", SCRIPT_DIR)
OUTPUT_DIR = os.environ.get("OUTPUT_FOLDER", os.path.join(SCRIPT_DIR, "merged_files"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Année dans le nom final
FINAL_YEAR = os.environ.get("FINAL_YEAR") or "2025"  # mets str(datetime.now().year) si tu préfères dynamique

# Colonnes à retirer du rendu final (feuille QUIZ)
DROP_COLUMNS = {"Numéro", "Nom", "Importante", "source_fichier"}

# Largeurs de colonnes (QUIZ)
COLUMN_WIDTHS = {
    "Question": 140,
    "Type de question": 24,
    "Réponse": 70,
    "Valide": 10,
    "Feedback": 140,
    "NbCar Feedback": 16,
}

# Colonnes sur lesquelles on fusionne visuellement les cellules si valeurs consécutives identiques
MERGE_VISUAL_COLS = ["Question", "Feedback"]  # tu peux ajouter "Réponse" si utile


# ============================================================
# UTILITAIRES
# ============================================================
def read_csv_robust(path: str) -> pd.DataFrame:
    """Lecture robuste CSV (séparateur/encodage auto), dtype=str pour préserver les zéros initiaux."""
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
    Remplace le reste par un espace, compresse les espaces.
    """
    title = re.sub(r"[^\w\s\(\)\-àâäéèêëîïôöùûüçÀÂÄÉÈÊËÎÏÔÖÙÛÜÇ]", " ", title)
    title = re.sub(r"\s+", " ", title).strip()
    return title


def build_final_name_from_content(df_first: pd.DataFrame) -> str:
    """Construit '3001_Antirétroviraux (1)_biblio_2025.xlsx' depuis les colonnes 'Numéro' et 'Nom'."""
    # Numéro
    numero = None
    if "Numéro" in df_first.columns:
        s = df_first["Numéro"].dropna().astype(str).str.strip()
        s = s.replace("", pd.NA).dropna()
        if not s.empty:
            numero = s.iloc[0]
    if not numero:
        numero = "0000"
    m = re.match(r"^(\d{4})", str(numero))
    numero_clean = m.group(1) if m else str(numero).strip()

    # Titre
    titre = None
    if "Nom" in df_first.columns:
        s = df_first["Nom"].dropna().astype(str).str.strip()
        s = s.replace("", pd.NA).dropna()
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
    s = s.strip()
    return s


def normalize_df_for_dedup(df: pd.DataFrame) -> pd.Series:
    """Construit une clé de déduplication normalisée sur toutes les colonnes du DF."""
    norm_cols = []
    for c in df.columns:
        col = df[c].fillna("").astype(str).map(_normalize_text)
        norm_cols.append(col)
    # concat sécurisée
    key = norm_cols[0]
    for i in range(1, len(norm_cols)):
        key = key + "||" + norm_cols[i]
    return key


def write_quiz_sheet(wb: Workbook, df: pd.DataFrame) -> None:
    """Crée/écrit la feuille QUIZ (données + format + formules + fusion visuelle)."""
    ws = wb.active
    ws.title = "QUIZ"

    # Écrire l'entête
    headers = list(df.columns)
    ws.append(headers)

    # Écrire les données
    for row in df.itertuples(index=False, name=None):
        ws.append(row)

    # Mise en forme
    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(ws.max_column)
    ws.auto_filter.ref = f"A1:{last_col_letter}1"

    # Largeurs
    header_to_idx = {name: i + 1 for i, name in enumerate(headers)}
    for col_name, width in COLUMN_WIDTHS.items():
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

    # Bordures fines
    thin = Border(
        left=Side(style="thin", color="CCCCCC"),
        right=Side(style="thin", color="CCCCCC"),
        top=Side(style="thin", color="CCCCCC"),
        bottom=Side(style="thin", color="CCCCCC"),
    )
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.border = thin

    # Fusion visuelle des cellules identiques consécutives (Question / Feedback)
    for col_name in MERGE_VISUAL_COLS:
        idx = header_to_idx.get(col_name)
        if not idx:
            continue
        start = 2
        prev_val = ws.cell(row=2, column=idx).value if ws.max_row >= 2 else None
        for r in range(3, ws.max_row + 1):
            val = ws.cell(row=r, column=idx).value
            if val != prev_val:
                # fusionner [start, r-1] si au moins 2 lignes
                if r - 1 > start:
                    ws.merge_cells(start_row=start, end_row=r - 1, start_column=idx, end_column=idx)
                    # aligner en haut
                    ws.cell(row=start, column=idx).alignment = Alignment(wrap_text=True, vertical="top")
                start = r
                prev_val = val
        # fin de colonne : fusionner la dernière séquence
        if ws.max_row >= 2 and ws.max_row > start:
            ws.merge_cells(start_row=start, end_row=ws.max_row, start_column=idx, end_column=idx)
            ws.cell(row=start, column=idx).alignment = Alignment(wrap_text=True, vertical="top")


def add_len_formula_on_feedback(wb: Workbook) -> None:
    """Ajoute la formule =LEN(Feedback) dans la colonne 'NbCar Feedback' déjà présente dans les colonnes du DataFrame."""
    ws = wb["QUIZ"]
    # Trouver index des colonnes
    headers = [c.value if c.value is not None else "" for c in next(ws.iter_rows(min_row=1, max_row=1))]
    header_to_idx = {name: i + 1 for i, name in enumerate(headers)}
    col_feedback = header_to_idx.get("Feedback")
    col_len = header_to_idx.get("NbCar Feedback")
    if not col_feedback or not col_len:
        return
    for r in range(2, ws.max_row + 1):
        # Utiliser la fonction anglaise (openpyxl), Excel FR affichera NBCAR
        ws.cell(row=r, column=col_len).value = f"=LEN({get_column_letter(col_feedback)}{r})"


# ============================================================
# PIPELINE PRINCIPAL
# ============================================================
def main():
    # 1) Récupérer les CSV à fusionner (2 fichiers attendus par usage, mais n>2 supporté)
    csv_files: List[str] = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".csv")]
    if not csv_files:
        print(f"Aucun CSV trouvé dans {INPUT_DIR}")
        return
    csv_files_sorted = sorted(csv_files)

    # 2) Lire le 1er CSV pour construire le nom final (à partir du CONTENU)
    df_first = read_csv_robust(os.path.join(INPUT_DIR, csv_files_sorted[0]))
    final_name = build_final_name_from_content(df_first)
    final_path = os.path.join(OUTPUT_DIR, final_name)

    # 3) Nettoyer d'anciens xlsx dans OUTPUT_DIR
    for name in os.listdir(OUTPUT_DIR):
        p = os.path.join(OUTPUT_DIR, name)
        if os.path.isfile(p) and name.lower().endswith(".xlsx"):
            try:
                os.remove(p)
            except Exception:
                pass

    # 4) Lire & concaténer tous les CSV fournis
    frames: List[pd.DataFrame] = [df_first]
    for name in csv_files_sorted[1:]:
        frames.append(read_csv_robust(os.path.join(INPUT_DIR, name)))
    fusion = pd.concat(frames, axis=0, ignore_index=True, sort=False)

    # 5) Supprimer colonnes inutiles
    keep_cols = [c for c in fusion.columns if c not in DROP_COLUMNS]
    fusion = fusion[keep_cols]

    # 6) Déduplication "intelligente" (normalisation)
    key = normalize_df_for_dedup(fusion)
    fusion = fusion.loc[~key.duplicated(keep="first")].reset_index(drop=True)

    # 7) Ajouter colonne NbCar Feedback (si Feedback existe)
    if "Feedback" in fusion.columns:
        if "NbCar Feedback" not in fusion.columns:
            fusion["NbCar Feedback"] = 0  # placeholder, les formules seront écrites ensuite

    # 8) Ordre de colonnes souhaité
    preferred = ["Question", "Type de question", "Réponse", "Valide", "Feedback", "NbCar Feedback"]
    ordered = [c for c in preferred if c in fusion.columns] + [c for c in fusion.columns if c not in preferred]
    fusion = fusion[ordered]

    # 9) Écrire le classeur final (QUIZ + BIBLIOGRAPHIE vide) + format
    wb = Workbook()
    write_quiz_sheet(wb, fusion)
    wb.create_sheet(title="BIBLIOGRAPHIE")

    # 10) Injecter les formules LEN() dans "NbCar Feedback"
    add_len_formula_on_feedback(wb)

    # 11) Sauvegarde finale
    wb.save(final_path)
    wb.close()

    # Expose le nom si un autre script l'utilise ensuite (optionnel)
    os.environ["FINAL_XLSX_NAME"] = final_name
    print(f"✅ Fichier final généré : {final_path}")

if __name__ == "__main__":
    main()