(function () {
  "use strict";

  function completeEvent(event) {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  }

  async function autoInsertSignature(event) {
    try {
      await window.TegeSignature.insertSignatureForCurrentFrom({ useDefault: false });
    } catch (error) {
      // Keep compose usable even when no matching signature exists.
      console.warn("TeGe signature auto insert skipped:", error.message);
    } finally {
      completeEvent(event);
    }
  }

  async function insertSignatureCommand(event) {
    try {
      await window.TegeSignature.insertSignatureForCurrentFrom({ useDefault: true });
    } catch (error) {
      console.error("TeGe signature insert failed:", error.message);
    } finally {
      completeEvent(event);
    }
  }

  Office.actions.associate("onNewMessageCompose", autoInsertSignature);
  Office.actions.associate("onMessageFromChanged", autoInsertSignature);
  Office.actions.associate("insertSignatureCommand", insertSignatureCommand);
})();
