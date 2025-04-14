
import os
import tempfile
import zipfile
import subprocess
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS

app = Flask(__name__)
CORS(app)
BASE_DIR = tempfile.mkdtemp()
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return "🚀 Backend Flask en ligne et opérationnel !"
@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        files = request.files.getlist('files')
        if not files:
            return jsonify({'error': 'Aucun fichier reçu'}), 400

        for file in files:
            print(f"Fichier reçu : {file.filename}")  # <-- log utile

        # (Optionnel) retourne un test pour voir si ça passe
        return jsonify({'message': 'Fichiers reçus avec succès'}), 200

    except Exception as e:
        print("Erreur dans /upload :", e)  # <-- trace dans Render logs
        return jsonify({'error': str(e)}), 500
        
@app.route('/upload', methods=['POST'])
def upload_files():
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'Aucun fichier reçu'}), 400

    # Nettoyage du dossier upload
    for f in os.listdir(UPLOAD_FOLDER):
        os.remove(os.path.join(UPLOAD_FOLDER, f))

    # Sauvegarde des fichiers CSV
    for file in files:
        file.save(os.path.join(UPLOAD_FOLDER, file.filename))

    try:
        # Exécution de Main.py en passant les chemins en arguments
        subprocess.run([
            'python', 'Main.py',
            UPLOAD_FOLDER,
            OUTPUT_FOLDER
        ], check=True)

        # Création du ZIP
        zip_path = os.path.join(BASE_DIR, 'resultats.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for root, dirs, files in os.walk(OUTPUT_FOLDER):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, OUTPUT_FOLDER)
                    zipf.write(full_path, arcname)

        return send_file(zip_path, as_attachment=True)

    except subprocess.CalledProcessError as e:
        return jsonify({'error': f'Erreur d’exécution : {e}'}), 500

import os

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)

