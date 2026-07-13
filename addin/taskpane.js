(function () {
  "use strict";

  function byId(id) {
    return document.getElementById(id);
  }

  function setStatus(message, isError) {
    const status = byId("status");
    status.textContent = message;
    status.style.color = isError ? "#9b1c1c" : "#4c665b";
  }

  function optionLabel(signature) {
    return `${signature.name} - ${signature.email}`;
  }

  async function initializeTaskpane() {
    const data = await window.TegeSignature.loadSignatureData();
    const fromEmail = await window.TegeSignature.getFromEmail();
    const select = byId("signatureSelect");

    data.signatures.forEach((signature) => {
      const option = document.createElement("option");
      option.value = signature.id;
      option.textContent = optionLabel(signature);
      option.setAttribute("data-email", signature.email);
      select.appendChild(option);
    });

    const match = window.TegeSignature.findSignature(data, fromEmail) || window.TegeSignature.findDefaultSignature(data);
    if (match) {
      select.value = match.id;
    }

    setStatus(`Absender erkannt: ${fromEmail || "unbekannt"}`);

    byId("insertButton").addEventListener("click", async () => {
      try {
        const signature = await window.TegeSignature.insertSignatureById(select.value, { useDefault: true });
        setStatus(`Eingefügt: ${signature.name}`);
      } catch (error) {
        setStatus(error.message, true);
      }
    });
  }

  Office.onReady(() => {
    initializeTaskpane().catch((error) => setStatus(error.message, true));
  });
})();
