import sys
import subprocess
import os

def main():
    if len(sys.argv) != 3:
        print("Utilisation : python Main.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_folder = sys.argv[2]

    # Passer les chemins comme variables d’environnement
    env = os.environ.copy()
    env["INPUT_FOLDER"] = input_folder
    env["OUTPUT_FOLDER"] = output_folder

    subprocess.run(['python', 'MergeCSV.py'], check=True, env=env)
    subprocess.run(['python', 'transformationxlsx.py'], check=True, env=env)

    print("✅ Traitement terminé")

if __name__ == "__main__":
    main()