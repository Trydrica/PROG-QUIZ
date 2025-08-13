# Imports à compléter en haut du fichier
import re
import csv
from io import BytesIO
import pandas as pd
from flask import Flask, request, send_file, jsonify

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        files = request.files.getlist('files')
        if not files:
            return jsonify({'error': 'Aucun fichier reçu'}), 400

        mem_zip = BytesIO()
        global_rows = []
        group_buckets = {}  # p.ex. {10: [df1, df2], 20: [df3], ...}

        def read_csv_robust(raw_bytes, filename):
            """Lecture robuste : délimiteur auto, encodage tolérant."""
            last_err = None
            for enc in ("utf-8-sig", "utf-8", "latin-1"):
                try:
                    sample = raw_bytes[:4096].decode(enc, errors='ignore')
                    try:
                        dialect = csv.Sniffer().sniff(sample)
                        sep = dialect.delimiter
                    except Exception:
                        sep = None  # laisse pandas inférer
                    return pd.read_csv(BytesIO(raw_bytes), sep=sep, engine='python', encoding=enc)
                except Exception as e:
                    last_err = e
            raise ValueError(f'Lecture CSV "{filename}" impossible: {last_err}')

        with zipfile.ZipFile(mem_zip, mode='w', compression=zipfile.ZIP_DEFLATED) as zf:
            for f in files:
                raw = f.read()
                df = read_csv_robust(raw, f.filename)

                # Ajoute la source pour tracer l’origine après fusion
                df.insert(0, 'source_fichier', f.filename)

                # 1) Écrire l’Excel individuel
                xbuf = BytesIO()
                with pd.ExcelWriter(xbuf, engine='openpyxl') as w:
                    df.to_excel(w, index=False, sheet_name='Données')
                xbuf.seek(0)
                indiv_name = os.path.splitext(f.filename)[0] + '.xlsx'
                zf.writestr(indiv_name, xbuf.getvalue())

                # 2) Alimente la fusion globale
                global_rows.append(df)

                # 3) Bucket par groupe 10/20/30… (extrait les 4 chiffres du nom)
                m = re.search(r'(\d{4})', f.filename)
                if m:
                    group = int(m.group(1)[:2])  # 1001 -> 10, 2003 -> 20
                    group_buckets.setdefault(group, []).append(df)

            # 4) Fusion globale (outer → on garde toutes les colonnes possibles)
            if global_rows:
                fusion_all = pd.concat(global_rows, axis=0, ignore_index=True, sort=False)
                xbuf = BytesIO()
                with pd.ExcelWriter(xbuf, engine='openpyxl') as w:
                    fusion_all.to_excel(w, index=False, sheet_name='Fusion')
                xbuf.seek(0)
                zf.writestr('fusion_globale.xlsx', xbuf.getvalue())

            # 5) Fusions par groupe (si applicables)
            for group, dfs in group_buckets.items():
                fusion_g = pd.concat(dfs, axis=0, ignore_index=True, sort=False)
                xbuf = BytesIO()
                with pd.ExcelWriter(xbuf, engine='openpyxl') as w:
                    fusion_g.to_excel(w, index=False, sheet_name=f'Groupe{group}')
                xbuf.seek(0)
                zf.writestr(f'Group_{group}.xlsx', xbuf.getvalue())

        mem_zip.seek(0)
        return send_file(mem_zip, as_attachment=True,
                         download_name='resultats.zip', mimetype='application/zip')

    except Exception as e:
        print("Erreur dans /upload :", e)
        return jsonify({'error': str(e)}), 500