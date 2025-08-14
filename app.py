#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import tempfile
import subprocess
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

# --- Chemins de base ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MERGE_SCRIPT = os.path.join(BASE_DIR, "MergeCSV.py")

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 Mo

# Autorise ton front (GitHub Pages) + localhost
CORS(app, resources={
    r"/upload": {
        "origins": [
            "https://trydrica.github.io",
            "https://*.github.io",
            "http://localhost",
            "http://localhost:*",
            "http://127.0.0.1:*"
        ]
    },
    r"/": {"origins": "*"}
})

@app.route("/", methods=["GET"])
def health():
    return "✅ Backend prêt (retourne un .xlsx directement).", 200

@app.route("/upload", methods=["POST"])
def upload_files():
    try:
        files = request.files.getlist("files")
        if not files:
            return jsonify({"error": "Aucun fichier reçu"}), 400

        # Crée des dossiers temporaires isolés par requête
        with tempfile.TemporaryDirectory() as in_dir, tempfile.TemporaryDirectory() as out_dir:
            # 1) Sauvegarde les CSV
            for f in files:
                name = f.filename or "file.csv"
                if not name.lower().endswith(".csv"):
                    return jsonify({"error": f"Format non supporté : {name} (CSV attendu)"}), 400
                f.save(os.path.join(in_dir, name))

            # 2) Exécute MergeCSV.py pour produire l'Excel final dans out_dir
            env = os.environ.copy()
            env["INPUT_FOLDER"] = in_dir
            env["OUTPUT_FOLDER"] = out_dir
            # Optionnel : fixer l'année dans le nom
            env.setdefault("FINAL_YEAR", "2025")

            proc = subprocess.run(
                [sys.executable, "-u", MERGE_SCRIPT],
                cwd=BASE_DIR,
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=600  # 10 min
            )

            if proc.returncode != 0:
                return jsonify({
                    "error": "Échec lors de l'exécution de MergeCSV.py",
                    "stdout": proc.stdout[-4000:],
                    "stderr": proc.stderr[-4000:]
                }), 500

            # 3) Récupère l'unique .xlsx produit et l'envoie directement
            xlsx_files = [f for f in os.listdir(out_dir) if f.lower().endswith(".xlsx")]
            if not xlsx_files:
                return jsonify({
                    "error": "Aucun fichier Excel généré",
                    "stdout": proc.stdout[-4000:],
                    "stderr": proc.stderr[-4000:]
                }), 500

            # S'il y en a plusieurs (improbable), on prend le premier trié
            xlsx_files.sort()
            final_name = xlsx_files[0]
            final_path = os.path.join(out_dir, final_name)

            # Renvoie DIRECTEMENT l'Excel (pas de ZIP)
            return send_file(
                final_path,
                as_attachment=True,
                download_name=final_name,
                mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except subprocess.TimeoutExpired:
        return jsonify({"error": "Timeout lors de l'exécution de MergeCSV.py (600s)."}), 504
    except Exception as e:
        print("Erreur /upload :", e)
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)