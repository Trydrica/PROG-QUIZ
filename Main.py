#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import subprocess
import os

def main():
    if len(sys.argv) != 3:
        print("Utilisation : python Main.py <input_folder> <output_folder>")
        sys.exit(1)

    input_folder = os.path.abspath(sys.argv[1])
    output_folder = os.path.abspath(sys.argv[2])

    if not os.path.isdir(input_folder):
        print(f"Erreur : input_folder introuvable : {input_folder}")
        sys.exit(2)

    os.makedirs(output_folder, exist_ok=True)

    # Passer les chemins aux scripts enfants
    env = os.environ.copy()
    env["INPUT_FOLDER"] = input_folder
    env["OUTPUT_FOLDER"] = output_folder

    base_dir = os.path.dirname(os.path.abspath(__file__))
    merge_csv_script = os.path.join(base_dir, "MergeCSV.py")
    transformation_script = os.path.join(base_dir, "transformationxlsx.py")

    print("➡️  Exécution de MergeCSV.py ...")
    subprocess.run([sys.executable, "-u", merge_csv_script], check=True, env=env, cwd=base_dir)

    print("➡️  Exécution de transformationxlsx.py ...")
    subprocess.run([sys.executable, "-u", transformation_script], check=True, env=env, cwd=base_dir)

    print("✅ Pipeline terminé")

if __name__ == "__main__":
    main()