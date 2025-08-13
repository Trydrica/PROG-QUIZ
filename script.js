// --- Remise à zéro & wiring au chargement de la page ---
document.addEventListener("DOMContentLoaded", () => {
  const fileInput = document.getElementById("csvFiles");
  const processBtn = document.getElementById("processBtn");
  const resultDiv = document.getElementById("result");

  // Crée un bloc d'info s'il n'existe pas (pour afficher "X fichiers sélectionnés.")
  let fileInfo = document.getElementById("fileInfo");
  if (!fileInfo) {
    fileInfo = document.createElement("div");
    fileInfo.id = "fileInfo";
    fileInfo.style.marginTop = "6px";
    fileInput.insertAdjacentElement("afterend", fileInfo);
  }

  // Fonction utilitaire: reset de l'UI
  const resetUI = () => {
    fileInfo.textContent = "Aucun fichier sélectionné.";
    fileInput.value = "";
    resultDiv.innerHTML = "";
  };

  // Reset à chaque refresh de page
  resetUI();

  // Met à jour le message quand l'utilisateur choisit des fichiers
  fileInput.addEventListener("change", () => {
    const n = fileInput.files?.length || 0;
    if (n === 0) {
      fileInfo.textContent = "Aucun fichier sélectionné.";
    } else if (n === 1) {
      fileInfo.textContent = `1 fichier sélectionné : ${fileInput.files[0].name}`;
    } else {
      fileInfo.textContent = `${n} fichiers sélectionnés.`;
    }
    // Nettoie le résultat précédent si on re-choisit des fichiers
    resultDiv.innerHTML = "";
  });

  // --- Traitement & téléchargement ---
  processBtn.addEventListener("click", async () => {
    const files = fileInput.files;
    if (!files || files.length === 0) {
      alert("Veuillez sélectionner un ou plusieurs fichiers CSV.");
      return;
    }

    processBtn.disabled = true;
    processBtn.textContent = "Traitement en cours…";
    resultDiv.innerHTML = "⏳ Traitement en cours...";

    const formData = new FormData();
    for (const f of files) formData.append("files", f);

    try {
      const response = await fetch("https://prog-quiz-bmxz.onrender.com/upload", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        // Essaie de lire une éventuelle erreur JSON, sinon texte brut
        let msg;
        try {
          const j = await response.json();
          msg = j?.error || response.statusText;
        } catch {
          msg = await response.text();
        }
        resultDiv.innerHTML = `❌ Erreur côté serveur : ${msg}`;
        return;
      }

      // Vérifie que c'est bien un ZIP
      const ct = response.headers.get("content-type") || "";
      if (!ct.includes("application/zip")) {
        const text = await response.text();
        resultDiv.innerHTML = `⚠️ Réponse inattendue (pas un ZIP) :<br><pre style="text-align:left;white-space:pre-wrap;">${text}</pre>`;
        return;
      }

      // Ok, on télécharge le ZIP
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = "resultats.zip";
      link.textContent = "📦 Télécharger les résultats (ZIP)";
      resultDiv.innerHTML = "";
      resultDiv.appendChild(link);
    } catch (err) {
      resultDiv.innerHTML = `❌ Une erreur est survenue : ${err}`;
    } finally {
      processBtn.disabled = false;
      processBtn.textContent = "Lancer le traitement";
    }
  });
});