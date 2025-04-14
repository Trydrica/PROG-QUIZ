
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
        
if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)