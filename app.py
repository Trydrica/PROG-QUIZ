#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import tempfile
import zipfile
from io import BytesIO
from datetime import datetime

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

# --- important : on importe le pipeline exact défini dans Main.py ---
# process_folder(input_dir, output_dir) -> (list_xlsx_individuels, list_xlsx_fusions)
from Main import process_folder

app = Flask(__name__)

# Limite (optionnelle) : 50 Mo par requête
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

# CORS : autorise ton front (GitHub Pages) + localhost pour tests
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
    return "🚀 Backend Flask en ligne et opérationnel !", 200


@app.route("/upload", methods=["POST"])
def upload_files():
    """
    Reçoit plusieurs CSV, lance le pipeline Main.process_folder(),
    puis renvoie un ZIP contenant :
      - tous les .xlsx individuels
      - fusion_globale.xlsx
      - Group_xx.xlsx (si applicable)
    """
    try:
        files = request.files.getlist("files")
        if not files:
            return jsonify({"error": "Aucun fichier reçu"}), 400

        # Répertoires temporaires isolés par requête
        with tempfile.TemporaryDirectory() as in_dir, tempfile.TemporaryDirectory() as out_dir:
            # 1) Sauvegarde des CSV reçus
            for f in files:
                # sécurité basique sur le nom et écriture
                fname = f.filename or "file.csv"
                dest = os.path.join(in_dir, fname)
                f.save(dest)

            # 2) Lancer le pipeline EXACT (Main.py)
            #    -> crée les .xlsx individuels + fusion_globale + Group_xx dans out_dir
            indiv_paths, fusion_paths = process_folder(in_dir, out_dir)

            # 3) Zipper tout ce qui est produit (individuels + fusions)
            mem_zip = BytesIO()
            with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                # Ajoute d'abord les individuels (tri pour déterminisme)
                for p in sorted(indiv_paths):
                    zf.write(p, arcname=os.path.basename(p))
                # Puis les fusions
                for p in sorted(fusion_paths):
                    zf.write(p, arcname=os.path.basename(p))

            mem_zip.seek(0)
            # Nom de zip horodaté (optionnel)
            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            return send_file(
                mem_zip,
                as_attachment=True,
                download_name=f"resultats-{ts}.zip",
                mimetype="application/zip",
            )

    except Exception as e:
        # Log minimal en console (visible sur Render)
        print("Erreur dans /upload :", e)
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    # Port imposé par l’hébergeur (Render) ou 5000 en local
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)