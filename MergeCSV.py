import os

input_folder = os.environ.get("INPUT_FOLDER")
output_folder = os.environ.get("OUTPUT_FOLDER")

print("📥 input_folder =", input_folder)
print("📤 output_folder =", output_folder)

import os
import pandas as pd
import re

# Définir le répertoire contenant les fichiers CSV
input_directory = os.path.dirname(os.path.abspath(__file__))
output_directory = os.path.join(input_directory, "merged_files")
os.makedirs(output_directory, exist_ok=True)

# Charger tous les fichiers CSV dans le dossier
all_files = [f for f in os.listdir(input_directory) if f.endswith('.csv')]

# Fonction pour extraire le préfixe numérique (10xx, 20xx, 30xx, etc.)
def extract_group_number(filename):
    match = re.search(r'\D*(\d{2})\d{2}', filename)
    return int(match.group(1)) if match else None

# Créer un dictionnaire pour regrouper les fichiers par préfixe numérique
files_by_group = {}

for file in all_files:
    group_number = extract_group_number(file)
    if group_number is not None:
        if group_number not in files_by_group:
            files_by_group[group_number] = []
        files_by_group[group_number].append(file)
    else:
        print(f"Fichier ignoré : {file} (pas de préfixe valide trouvé)")

# Fusionner les fichiers pour chaque groupe numérique
for group, files in files_by_group.items():
    merged_df = pd.DataFrame()
    sorted_files = sorted(files, key=lambda x: int(re.search(r'(\d+)', x).group(1)))

    print(f"Fusion des fichiers pour le groupe {group} : {files}")
    for file in sorted_files:
        file_path = os.path.join(input_directory, file)
        try:
            df = pd.read_csv(file_path)
            merged_df = pd.concat([merged_df, df], ignore_index=True)
        except Exception as e:
            print(f"Erreur lors de la lecture de {file}: {e}")

    if not merged_df.empty:
        try:
            # Sauvegarder chaque groupe dans un fichier Excel séparé
            output_file = os.path.join(output_directory, f"Group_{group}.xlsx")
            merged_df.to_excel(output_file, index=False)
            print(f"Fichier créé : {output_file}")
        except Exception as e:
            print(f"Erreur lors de la sauvegarde du fichier pour le groupe {group}: {e}")
    else:
        print(f"Aucune donnée fusionnée pour le groupe {group}")

import os
output_file = "/Users/romainpoulin/.../Group_20.xlsx"
if not os.path.exists(output_file):
    print(f"Erreur : le fichier {output_file} n'a pas été créé.")
else:
    print(f"Le fichier {output_file} a bien été créé.")

print(f"Les fichiers CSV ont été fusionnés et enregistrés dans le dossier : {output_directory}")
