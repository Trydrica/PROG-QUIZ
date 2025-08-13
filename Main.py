#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import shutil
import subprocess
from flask import Flask, request, send_file, jsonify

# Répertoires
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FOLDER = os.path.join(BASE_DIR, "input_files")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "merged_files")
MERGE_SCRIPT = os.path.join(BASE_DIR, "MergeCSV.py")

# Création des dossiers
os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)

def clean_folders():
    """Vide les dossiers INPUT_FOLDER et OUTPUT_FOLDER."""
    for folder in (INPUT_FOLDER, OUTPUT_FOLDER):
        for f in os.listdir(folder):
            path = os.path.join(folder, f)
            if os.path.isfile(path):
                os.remove(path)
            elif os.path.isdir(path):
                shutil.rmtree(path)

@app.route("/upload", methods=["POST"])
def upload_files():
    try:
        clean_folders()

        # Sauvegarde des fichiers CSV uploadés
        files = request.files.getlist("files")
        if not files:
            return jsonify({"error": "Aucun fichier reçu"}), 400

        for file in files:
            filename = file.filename
            if not filename.lower().endswith(".csv"):
                return jsonify({"error": f"Format non supporté : {filename}"}), 400
            file.save(os.path.join(INPUT_FOLDER, filename))

        # Exécute MergeCSV.py
        env = os.environ.copy()
        env["INPUT_FOLDER"] = INPUT_FOLDER
        env["OUTPUT_FOLDER"] = OUTPUT_FOLDER
        subprocess.run(["python", MERGE_SCRIPT], check=True, env=env)

        # Récupère le fichier final généré
        xlsx_files = [f for f in os.listdir(OUTPUT_FOLDER) if f.lower().endswith(".xlsx")]
        if not xlsx_files:
            return jsonify({"error": "Aucun fichier Excel généré"}), 500

        final_path = os.path.join(OUTPUT_FOLDER, xlsx_files[0])
        return send_file(final_path, as_attachment=True)

    except subprocess.CalledProcessError as e:
        return jsonify({"error": f"Erreur lors de l'exécution de MergeCSV.py: {e}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route("/", methods=["GET"])
def index():
    return "Serveur prêt à recevoir des CSV."

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)