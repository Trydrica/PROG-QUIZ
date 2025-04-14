from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import os
import tempfile
import zipfile
import subprocess

app = Flask(__name__)
CORS(app)

BASE_DIR = tempfile.mkdtemp()
UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "output")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def home():
    return "🎉 Backend Flask en ligne sur Railway !"

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        files = request.files.getlist('files')
        if not files:
            return jsonify({'error': 'Aucun fichier reçu'}), 400

        for file in files:
            file.save(os.path.join(UPLOAD_FOLDER, file.filename))

        print("📂 UPLOAD_FOLDER:", UPLOAD_FOLDER)
        print("📄 Fichiers présents après upload :", os.listdir(UPLOAD_FOLDER))

        env = os.environ.copy()
        env["INPUT_FOLDER"] = UPLOAD_FOLDER
        env["OUTPUT_FOLDER"] = OUTPUT_FOLDER

        subprocess.run(['python', 'Main.py', UPLOAD_FOLDER, OUTPUT_FOLDER], check=True, env=env)

        print("🧾 Contenu OUTPUT_FOLDER :", os.listdir(OUTPUT_FOLDER))

        if not os.listdir(OUTPUT_FOLDER):
            return jsonify({'error': 'Aucun fichier généré'}), 500

        zip_path = os.path.join(BASE_DIR, 'resultats.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for root, dirs, files in os.walk(OUTPUT_FOLDER):
                for file in files:
                    full_path = os.path.join(root, file)
                    arcname = os.path.relpath(full_path, OUTPUT_FOLDER)
                    zipf.write(full_path, arcname)

        return send_file(zip_path, as_attachment=True)

    except Exception as e:
        print("❌ Erreur dans /upload :", e)
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
