import sys
import subprocess
import os

def main():
    if len(sys.argv) != 3:
        print("Utilisation : python Main.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_folder = sys.argv[2]

    # Définir les variables d’environnement à passer aux autres scripts
    env = os.environ.copy()
    env["INPUT_FOLDER"] = input_folder
    env["OUTPUT_FOLDER"] = output_folder

    # Appeler MergeCSV.py avec les mêmes variables d’environnement
    subprocess.run(['python', 'MergeCSV.py'], check=True, env=env)

    # Puis transformationxlsx_old.py
    subprocess.run(['python', 'transformationxlsx_old.py'], check=True, env=env)

    print("Traitement terminé.")

if __name__ == "__main__":
    main()