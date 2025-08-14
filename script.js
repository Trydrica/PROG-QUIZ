document.addEventListener("DOMContentLoaded", () => {
  // --- Résolution "robuste" des éléments (essaie plusieurs IDs possibles) ---
  const pick = (ids) => ids.map(id => document.getElementById(id)).find(el => !!el) || null;

  const form      = pick(["uploadForm", "form-upload", "form"]);
  const fileInput = pick(["fileInput", "csvFiles", "file-input", "files"]);
  let   resultDiv = pick(["result", "output", "messages"]);
  let   button    = pick(["processBtn", "submitBtn", "runBtn"]);

  // Crée un conteneur d’état si absent
  if (!resultDiv) {
    resultDiv = document.createElement("div");
    resultDiv.id = "result";
    (form || document.body).appendChild(resultDiv);
  }

  // Si pas de bouton dédié, on utilisera la soumission du formulaire
  const useSubmitOnForm = !button && !!form;

  // Reset UI au chargement
  if (fileInput) fileInput.value = "";
  resultDiv.textContent = "";

  // Petite fonction d’affichage
  const show = (html) => { resultDiv.innerHTML = html; };

  // Handler principal
  const run = async () => {
    try {
      // Empêche double clic
      if (button) { button.disabled = true; button.dataset.originalText = button.textContent; button.textContent = "Traitement…"; }
      show("⏳ Traitement en cours...");

      // Vérifs de base
      if (!fileInput) {
        show("❌ Impossible de trouver le champ de fichiers sur la page (IDs attendus : fileInput / csvFiles / file-input / files).");
        return;
      }
      const files = fileInput.files;
      if (!files || !files.length) {
        show("⚠️ Veuillez sélectionner au moins un fichier CSV.");
        return;
      }

      // Prépare la requête
      const formData = new FormData();
      for (let i = 0; i < files.length; i++) {
        formData.append("files", files[i]);
      }

      // Appel backend (adapter l’URL si besoin)
      const response = await fetch("https://prog-quiz-bmxz.onrender.com/upload", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        let msg;
        try { msg = (await response.json()).error; } catch { msg = await response.text(); }
        show(`❌ Erreur serveur : ${msg || response.statusText}`);
        return;
      }

      // On attend un XLSX direct
      const ct = (response.headers.get("content-type") || "").toLowerCase();
      if (!ct.includes("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
        const text = await response.text();
        show(`⚠️ Réponse inattendue (pas un XLSX) :<br><pre style="white-space:pre-wrap;text-align:left;">${text}</pre>`);
        return;
      }

      // Récupère le nom proposé par le serveur
      const cd = response.headers.get("content-disposition") || "";
      let filename = "resultat.xlsx";
      const m = cd.match(/filename\*?=(?:UTF-8'')?"?([^\";]+)"?/i);
      if (m && m[1]) filename = decodeURIComponent(m[1]);

      // Téléchargement automatique
      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 100);

      show("✅ Téléchargement lancé.");
    } catch (err) {
      console.error(err);
      show(`❌ Erreur : ${err?.message || err}`);
    } finally {
      if (button) { button.disabled = false; button.textContent = button.dataset.originalText || "Lancer le traitement"; }
    }
  };

  // Wiring des événements
  if (useSubmitOnForm) {
    form.addEventListener("submit", (e) => { e.preventDefault(); run(); });
  } else if (button) {
    // Si un bouton dédié existe, on empêche la soumission classique du formulaire
    if (form) form.addEventListener("submit", (e) => e.preventDefault());
    button.addEventListener("click", (e) => { e.preventDefault(); run(); });
  } else {
    // Aucun form ni bouton détecté → on affiche une aide
    show("ℹ️ Aucun bouton ni formulaire détecté. Ajoute un bouton avec id <code>processBtn</code> ou un formulaire avec id <code>uploadForm</code>.");
  }
});