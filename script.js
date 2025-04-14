document.getElementById('processBtn').addEventListener('click', async () => {
    const files = document.getElementById('csvFiles').files;
    if (files.length === 0) {
      alert("Veuillez sélectionner un ou plusieurs fichiers CSV.");
      return;
    }
  
    const formData = new FormData();
    for (const file of files) {
      formData.append('files', file);
    }
  
    const response = await fetch('https://csv-to-excel-backend.onrender.com/upload', {      body: formData
    });
  
    if (response.ok) {
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = 'resultats.zip';
      link.textContent = 'Télécharger les fichiers traités (ZIP)';
      document.getElementById('result').innerHTML = '';
      document.getElementById('result').appendChild(link);
    } else {
      alert("Erreur lors du traitement.");
    }
  });