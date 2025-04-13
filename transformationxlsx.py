import os
input_folder = os.environ.get('INPUT_FOLDER')
output_folder = os.environ.get('OUTPUT_FOLDER')

import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from datetime import datetime

# Chemin du dossier contenant les fichiers merged_files
directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "merged_files")

# Vérifier si le dossier existe
if not os.path.exists(directory):
    raise FileNotFoundError(f"Le dossier '{directory}' n'existe pas.")

# Liste des fichiers Excel dans le dossier
file_list = [file for file in os.listdir(directory) if file.endswith('.xlsx')]

if not file_list:
    raise FileNotFoundError("Aucun fichier Excel trouvé dans le dossier 'merged_files'.")

# Traiter chaque fichier un par un
for file_name in file_list:
    input_file = os.path.join(directory, file_name)

    # Vérifier que le fichier existe et n'est pas vide
    if not os.path.exists(input_file):
        print(f"Fichier introuvable : {input_file}, ignoré.")
        continue

    if os.path.getsize(input_file) < 1000:
        print(f"Fichier suspect (trop léger ou vide) : {input_file}, ignoré.")
        continue

    # Charger le fichier Excel avec gestion des erreurs
    try:
        wb = load_workbook(input_file)
    except Exception as e:
        print(f"Erreur lors de l'ouverture de {input_file} : {e}")
        continue
    wb = load_workbook(input_file)
    ws = wb.active

    # Renommer la première feuille en "QUIZ"
    ws.title = "QUIZ"

    # Créer une nouvelle feuille appelée "BIBLIOGRAPHIE"
    biblio_sheet = wb.create_sheet(title="BIBLIOGRAPHIE")

    # Appliquer l'alignement dans la feuille "BIBLIOGRAPHIE"
    default_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    for row in biblio_sheet.iter_rows():
        for cell in row:
            cell.alignment = default_alignment

    # Définir le mapping des en-têtes
    header_map = {ws.cell(row=1, column=col).value: col for col in range(1, ws.max_column + 1)}

    # Étape 7 : Ajouter la formule NB.CAR() dans la dernière colonne pour compter les caractères de chaque cellule de la colonne Feedback
    if "Feedback" in header_map:
        feedback_col_index = header_map["Feedback"]
        new_col_index = ws.max_column + 1
        ws.cell(row=1, column=new_col_index, value="NbCar Feedback")
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            feedback_cell = ws.cell(row=row[0].row, column=feedback_col_index).coordinate
            formula = f"=LEN({feedback_cell})"
            ws.cell(row=row[0].row, column=new_col_index, value=formula)

    # Étape 1 : Fusionner toutes les cellules identiques sur les colonnes spécifiées
    columns_to_merge = ["Numéro", "Nom", "Question", "Feedback"]
    columns_to_merge_indices = [header_map[col_name] for col_name in columns_to_merge if col_name in header_map]

    for col in columns_to_merge_indices:
        merge_start = None
        previous_value = None
        for row_cells in ws.iter_rows(min_col=col, max_col=col):
            row = row_cells[0].row
            current_value = ws.cell(row=row, column=col).value
            if current_value == previous_value:
                if merge_start is None:
                    merge_start = row - 1
            else:
                if merge_start is not None:
                    ws.merge_cells(start_row=merge_start, start_column=col, end_row=row - 1, end_column=col)
                    ws.cell(row=merge_start, column=col).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
                    merge_start = None
            previous_value = current_value
        if merge_start is not None:
            ws.merge_cells(start_row=merge_start, start_column=col, end_row=ws.max_row, end_column=col)
            ws.cell(row=merge_start, column=col).alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Étape 2 : Ajouter des bordures noires et appliquer l'alignement à toutes les cellules non vides (y compris fusionnées)
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Appliquer les bordures et l'alignement aux cellules fusionnées
    for merged_range in ws.merged_cells.ranges:
        for row in ws[merged_range.coord]:
            for cell in row:
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Étape 3 : Renommer le fichier en utilisant la colonne Numéro et Nom
    numero_index = header_map.get("Numéro")
    nom_index = header_map.get("Nom")

    if numero_index and nom_index:
        used_rows = [row[0].row for row in ws.iter_rows(min_col=numero_index, max_col=numero_index) if row[0].value]
        for row in used_rows[1:]:
            numero_value = ws.cell(row=row, column=numero_index).value
            nom_value = ws.cell(row=row, column=nom_index).value
            if numero_value and nom_value:
                break
        else:
            raise ValueError("Aucune valeur valide trouvée dans les colonnes 'Numéro' et 'Nom'.")
    else:
        raise ValueError("Les colonnes 'Numéro' et 'Nom' sont nécessaires pour renommer le fichier.")

    current_year = datetime.now().year
    new_filename = f"{numero_value}_{nom_value}_biblio_{current_year}.xlsx"
    new_file_path = os.path.join(directory, new_filename)

    # Étape 4 : Masquer les colonnes "Numéro", "Nom", "Type de question" et "Importante"
    columns_to_hide = ["Numéro", "Nom", "Type de question", "Importante"]
    columns_to_hide_indices = [header_map[col_name] for col_name in columns_to_hide if col_name in header_map]
    for col_index in columns_to_hide_indices:
        ws.column_dimensions[ws.cell(row=1, column=col_index).column_letter].hidden = True

    # Étape 5 : Appliquer une hauteur de ligne de 150 pixels de la ligne 2 à la dernière ligne non vide et formatage
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        ws.row_dimensions[row[0].row].height = 150
        for cell in row:
            if cell.value is not None:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)

    # Étape 6 : Ajuster la largeur des colonnes spécifiées
    column_widths = {
        "Question": 50,  # Ajusté pour une largeur raisonnable en Excel
        "Réponse": 30,
        "Valide": 10,
        "Feedback": 140
    }
    for col_name, width in column_widths.items():
        if col_name in header_map:
            ws.column_dimensions[ws.cell(row=1, column=header_map[col_name]).column_letter].width = width

    # Sauvegarder le fichier modifié
    wb.save(new_file_path)
    wb.close()

    print(f"Le fichier a été modifié et enregistré sous : {new_file_path}")

print("Tous les fichiers ont été traités.")
