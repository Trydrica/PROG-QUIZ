document.getElementById('processBtn').addEventListener('click', async () => {
  const files = document.getElementById('csvFiles').files;
  const resultDiv = document.getElementById('result');

  if (files.length === 0) {
    alert("Veuillez sélectionner un ou plusieurs fichiers CSV.");
    return;
  }

  resultDiv.innerHTML = '⏳ Traitement en cours...';

  const formData = new FormData();
  for (const file of files) {
    formData.append('files', file);
  }

  try {
    const response = await fetch('https://prog-quiz-bmxz.onrender.com/upload', {
      method: 'POST',
      body: formData
    });

    if (!response.ok) {
      const error = await response.json();
      resultDiv.innerHTML = `❌ Erreur côté serveur : ${error.error || response.statusText}`;
      return;
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'resultats.zip';
    link.textContent = '📦 Télécharger les résultats (ZIP)';
    resultDiv.innerHTML = '';
    resultDiv.appendChild(link);
  } catch (err) {
    resultDiv.innerHTML = `❌ Une erreur est survenue : ${err}`;
  }
});