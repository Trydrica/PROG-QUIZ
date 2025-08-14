document.addEventListener("DOMContentLoaded", () => {
  const uploadForm = document.getElementById("uploadForm");
  const fileInput = document.getElementById("fileInput");
  const resultDiv = document.getElementById("result");

  // Réinitialise l'affichage au refresh
  if (fileInput) {
    fileInput.value = "";
  }

  uploadForm.addEventListener("submit", async (e) => {
    e.preventDefault();

    const files = fileInput.files;
    if (!files.length) {
      resultDiv.innerHTML = "⚠️ Veuillez sélectionner au moins un fichier CSV.";
      return;
    }

    const formData = new FormData();
    for (let i = 0; i < files.length; i++) {
      formData.append("files", files[i]);
    }

    resultDiv.innerHTML = "⏳ Traitement en cours...";

    try {
      const response = await fetch("https://prog-quiz-bmxz.onrender.com/upload", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let msg;
        try {
          msg = (await response.json()).error;
        } catch {
          msg = await response.text();
        }
        resultDiv.innerHTML = `❌ Erreur serveur : ${msg || response.statusText}`;
        return;
      }

      // Vérifie que la réponse est bien un Excel
      const ct = response.headers.get("content-type") || "";
      if (!ct.includes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
        const text = await response.text();
        resultDiv.innerHTML = `⚠️ Réponse inattendue (pas un XLSX) :<br><pre>${text}</pre>`;
        return;
      }

      // Récupère le nom depuis Content-Disposition
      const cd = response.headers.get("content-disposition") || "";
      let filename = "resultat.xlsx";
      const match = cd.match(/filename="?([^"]+)"?/i);
      if (match && match[1]) filename = match[1];

      // Création du blob et téléchargement automatique
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);

      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click(); // Lance le téléchargement
      a.remove();

      // Libère l'URL objet après un petit délai
      setTimeout(() => URL.revokeObjectURL(url), 100);

      resultDiv.innerHTML = "✅ Téléchargement lancé.";
    } catch (error) {
      console.error(error);
      resultDiv.innerHTML = `❌ Erreur : ${error.message}`;
    }
  });
});