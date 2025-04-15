#!/bin/bash

echo "🔁 Remplacement des fichiers corrigés..."
cp MergeCSV_final2.py MergeCSV.py
cp transformationxlsx_final2.py transformationxlsx.py

echo "🧹 Nettoyage des fichiers compilés..."
rm -rf __pycache__
find . -name "*.pyc" -delete

echo "📦 Ajout au commit..."
git add MergeCSV.py transformationxlsx.py

echo "📝 Commit forcé..."
git commit -m 'Force overwrite des scripts corrigés (chemins output_folder)'

echo "🚀 Push vers GitHub..."
git push

echo "✅ Terminé ! Attends le déploiement Railway puis teste un upload."
