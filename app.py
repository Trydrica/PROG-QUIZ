import os
import tempfile
import zipfile
import pandas as pd
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

        # Nettoyage des anciens fichiers
        for f in os.listdir(UPLOAD_FOLDER):
            os.remove(os.path.join(UPLOAD_FOLDER, f))
        for f in os.listdir(OUTPUT_FOLDER):
            os.remove(os.path.join(OUTPUT_FOLDER, f))

        # Sauvegarde et conversion
        for file in files:
            filename = file.filename
            file_path = os.path.join(UPLOAD_FOLDER, filename)
            file.save(file_path)

            df = pd.read_csv(file_path)
            excel_name = filename.replace('.csv', '.xlsx')
            excel_path = os.path.join(OUTPUT_FOLDER, excel_name)
            df.to_excel(excel_path, index=False)

        # Création du ZIP
        zip_path = os.path.join(BASE_DIR, 'resultats.zip')
        with zipfile.ZipFile(zip_path, 'w') as zipf:
            for f in os.listdir(OUTPUT_FOLDER):
                zipf.write(os.path.join(OUTPUT_FOLDER, f), f)

        return send_file(zip_path, as_attachment=True, download_name='resultats.zip', mimetype='application/zip')

    except Exception as e:
        print("Erreur dans /upload :", e)
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)