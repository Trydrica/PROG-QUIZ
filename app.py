#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import sys
import tempfile
import zipfile
from io import BytesIO
from datetime import datetime
import subprocess

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

app = Flask(__name__)
# Limite (optionnelle) : 50 Mo par requête
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

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
    return "🚀 Backend Flask en ligne et opérationnel !", 200


@app.route("/upload", methods=["POST"])
def upload_files():
    """
    Reçoit plusieurs CSV, exécute Main.py <input_dir> <output_dir>,
    puis renvoie un ZIP contenant exactement les fichiers produits par Main.py
    (individuels + fusion_globale + Group_xx + toute mise en forme spécifique).
    """
    try:
        files = request.files.getlist("files")
        if not files:
            return jsonify({"error": "Aucun fichier reçu"}), 400

        # Dossiers temporaires isolés par requête
        with tempfile.TemporaryDirectory() as in_dir, tempfile.TemporaryDirectory() as out_dir:
            # 1) Sauvegarde des CSV uploadés
            for f in files:
                fname = f.filename or "file.csv"
                dest = os.path.join(in_dir, fname)
                f.save(dest)

            # 2) Lancer TON pipeline local identique : Main.py <in_dir> <out_dir>
            #    - sys.executable garantit le bon interpréteur (Render/venv)
            #    - cwd=repo root pour que Main.py retrouve ses imports/scripts voisins
            cmd = [sys.executable, "-u", "Main.py", in_dir, out_dir]
            proc = subprocess.run(
                cmd,
                cwd=os.getcwd(),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                timeout=600  # 10 min de marge si gros fichiers
            )

            if proc.returncode != 0:
                # On renvoie les logs pour debug côté front si besoin
                return jsonify({
                    "error": "Échec lors de l'exécution de Main.py",
                    "stdout": proc.stdout[-4000:],  # tronqué pour éviter la surcharge
                    "stderr": proc.stderr[-4000:]
                }), 500

            # 3) Zipper tout ce que Main.py a produit dans out_dir
            produced = [f for f in os.listdir(out_dir) if os.path.isfile(os.path.join(out_dir, f))]
            if not produced:
                return jsonify({
                    "error": "Aucun fichier de sortie généré par Main.py",
                    "stdout": proc.stdout[-4000:],
                    "stderr": proc.stderr[-4000:]
                }), 500

            mem_zip = BytesIO()
            with zipfile.ZipFile(mem_zip, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                for name in sorted(produced):
                    full = os.path.join(out_dir, name)
                    zf.write(full, arcname=name)

            mem_zip.seek(0)
            ts = datetime.now().strftime("%Y%m%d-%H%M%S")
            return send_file(
                mem_zip,
                as_attachment=True,
                download_name=f"resultats-{ts}.zip",
                mimetype="application/zip",
            )

    except subprocess.TimeoutExpired:
        return jsonify({"error": "Timeout lors de l'exécution de Main.py (600s)."}), 504
    except Exception as e:
        print("Erreur dans /upload :", e)
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    # Port imposé par l’hébergeur (Render) ou 5000 en local
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)