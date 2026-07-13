(function () {
  "use strict";

  const SIGNATURE_DATA_URL = "https://jorgecastelan-svg.github.io/tege-signature-assets/addin/signatures.json";

  let cachedData = null;

  function normalizeEmail(value) {
    return String(value || "").trim().toLowerCase();
  }

  function getMailboxEmail() {
    try {
      return normalizeEmail(Office.context.mailbox.userProfile.emailAddress);
    } catch (error) {
      return "";
    }
  }

  function getFromEmail() {
    return new Promise((resolve) => {
      const fallback = getMailboxEmail();
      const item = Office.context.mailbox && Office.context.mailbox.item;

      if (!item || !item.from || typeof item.from.getAsync !== "function") {
        resolve(fallback);
        return;
      }

      item.from.getAsync((result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
          resolve(fallback);
          return;
        }

        resolve(normalizeEmail(result.value.emailAddress || result.value.displayName || fallback));
      });
    });
  }

  async function loadSignatureData() {
    if (cachedData) {
      return cachedData;
    }

    const response = await fetch(SIGNATURE_DATA_URL, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`Signaturdaten konnten nicht geladen werden (${response.status}).`);
    }

    cachedData = await response.json();
    return cachedData;
  }

  function findSignature(data, email) {
    const normalized = normalizeEmail(email);
    return data.signatures.find((item) => normalizeEmail(item.email) === normalized) || null;
  }

  function findSignatureById(data, id) {
    const normalized = String(id || "").trim().toLowerCase();
    return data.signatures.find((item) => String(item.id || "").trim().toLowerCase() === normalized) || null;
  }

  function findDefaultSignature(data) {
    return findSignature(data, data.defaultEmail) || data.signatures[0] || null;
  }

  function setSignatureHtml(html) {
    return new Promise((resolve, reject) => {
      const body = Office.context.mailbox.item.body;
      const options = { coercionType: Office.CoercionType.Html };

      if (typeof body.setSignatureAsync === "function") {
        body.setSignatureAsync(html, options, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error(result.error && result.error.message ? result.error.message : "setSignatureAsync failed."));
          }
        });
        return;
      }

      body.setSelectedDataAsync(html, options, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error && result.error.message ? result.error.message : "setSelectedDataAsync failed."));
        }
      });
    });
  }

  async function insertSignatureByEmail(email, options) {
    const data = await loadSignatureData();
    const signature = findSignature(data, email) || (options && options.useDefault ? findDefaultSignature(data) : null);

    if (!signature) {
      throw new Error(`Keine Signatur für ${email || "diese Mailbox"} gefunden.`);
    }

    await setSignatureHtml(signature.html);
    return signature;
  }

  async function insertSignatureById(id, options) {
    const data = await loadSignatureData();
    const signature = findSignatureById(data, id) || (options && options.useDefault ? findDefaultSignature(data) : null);

    if (!signature) {
      throw new Error(`Keine Signatur für ${id || "diese Auswahl"} gefunden.`);
    }

    await setSignatureHtml(signature.html);
    return signature;
  }

  async function insertSignatureForCurrentFrom(options) {
    const email = await getFromEmail();
    return insertSignatureByEmail(email, options);
  }

  window.TegeSignature = {
    getFromEmail,
    getMailboxEmail,
    loadSignatureData,
    findSignature,
    findSignatureById,
    findDefaultSignature,
    insertSignatureByEmail,
    insertSignatureById,
    insertSignatureForCurrentFrom,
  };
})();
