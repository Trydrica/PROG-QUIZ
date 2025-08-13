#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import csv
from io import BytesIO
from typing import Optional, List

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# -------------------------------------------------------------------
# RÉPERTOIRES D’ENTRÉE/SORTIE
# -------------------------------------------------------------------
# Si Main.py a défini INPUT_FOLDER / OUTPUT_FOLDER, on les utilise.
# Sinon, on retombe sur l’ancien comportement : dossier du script et "merged_files".
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.environ.get("INPUT_FOLDER", SCRIPT_DIR)
OUTPUT_DIR = os.environ.get("OUTPUT_FOLDER", os.path.join(SCRIPT_DIR, "merged_files"))
os.makedirs(OUTPUT_DIR, exist_ok=True)

# -------------------------------------------------------------------
# UTILITAIRES
# -------------------------------------------------------------------
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

def extract_group_from_filename(filename: str) -> Optional[int]:
    """
    Extrait un groupe 10/20/30… à partir d'un code 4 chiffres dans le nom.
    ex. quiz-1001.csv -> 10 ; abc_2033.csv -> 20 ; sinon None.
    """
    m = re.search(r"(\d{4})", filename)
    return int(m.group(1)[:2]) if m else None

def write_xlsx_with_format(df: pd.DataFrame, out_path: str, sheet_name: str = "Données") -> None:
    """
    Écrit un DataFrame en .xlsx + quelques mises en forme utiles :
      - Freeze en-tête (A2)
      - Auto-filtre sur la ligne d'entête
      - Largeurs de colonnes usuelles (si les titres existent)
      - Wrap text pour colonnes longues (Question/Réponse/Feedback)
    """
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet_name)

    wb = load_workbook(out_path)
    ws = wb[sheet_name]

    # Figer l’entête + filtre
    ws.freeze_panes = "A2"
    last_col_letter = get_column_letter(ws.max_column)
    ws.auto_filter.ref = f"A1:{last_col_letter}1"

    # Largeurs courantes (adapter si besoin à tes entêtes exactes)
    column_widths = {
        "Numéro": 10,
        "Nom": 30,
        "Question": 140,
        "Type de question": 24,
        "Réponse": 70,
        "Valide": 10,
        "Importante": 14,
        "Feedback": 140,
    }
    headers = [c.value if c.value is not None else "" for c in ws[1]]
    header_to_idx = {str(name): i + 1 for i, name in enumerate(headers)}
    for col_name, width in column_widths.items():
        idx = header_to_idx.get(col_name)
        if idx:
            ws.column_dimensions[get_column_letter(idx)].width = width

    # Wrap text sur colonnes longues
    wrap_cols = {"Question", "Réponse", "Feedback"}
    for col_name in wrap_cols:
        idx = header_to_idx.get(col_name)
        if idx:
            for row in ws.iter_rows(min_row=2, min_col=idx, max_col=idx):
                for cell in row:
                    cell.alignment = Alignment(wrap_text=True, vertical="top")

    wb.save(out_path)
    wb.close()

# -------------------------------------------------------------------
# TRAITEMENT
# -------------------------------------------------------------------
def main() -> None:
    # 1) lister les CSV à traiter
    csv_files: List[str] = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".csv")]
    if not csv_files:
        print(f"Aucun .csv trouvé dans : {INPUT_DIR}")
        return

    global_dfs: List[pd.DataFrame] = []
    per_group: dict[int, List[pd.DataFrame]] = {}

    # 2) pour chaque CSV : lire, ajouter une colonne source, écrire un .xlsx individuel
    for name in sorted(csv_files):
        src_path = os.path.join(INPUT_DIR, name)
        df = read_csv_robust(src_path)

        # garde la trace de la source
        df.insert(0, "source_fichier", name)

        # Excel individuel
        base = os.path.splitext(name)[0]
        out_indiv = os.path.join(OUTPUT_DIR, f"{base}.xlsx")
        write_xlsx_with_format(df, out_indiv, sheet_name="Données")
        print(f"✅ Fichier Excel créé : {out_indiv}")

        global_dfs.append(df)

        grp = extract_group_from_filename(name)
        if grp is not None:
            per_group.setdefault(grp, []).append(df)

    # 3) fusion globale
    if global_dfs:
        fusion_all = pd.concat(global_dfs, axis=0, ignore_index=True, sort=False)
        out_fusion = os.path.join(OUTPUT_DIR, "fusion_globale.xlsx")
        write_xlsx_with_format(fusion_all, out_fusion, sheet_name="Fusion")
        print(f"✅ Fusion globale : {out_fusion}")

    # 4) fusions par groupe (10/20/30…)
    for grp, dfs in sorted(per_group.items()):
        fusion_g = pd.concat(dfs, axis=0, ignore_index=True, sort=False)
        out_grp = os.path.join(OUTPUT_DIR, f"Group_{grp}.xlsx")
        write_xlsx_with_format(fusion_g, out_grp, sheet_name=f"Groupe{grp:02d}")
        print(f"✅ Fusion groupe {grp:02d} : {out_grp}")

    print(f"\nTerminé. Fichiers disponibles dans : {OUTPUT_DIR}")

if __name__ == "__main__":
    main()