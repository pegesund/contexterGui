// NorskTale Google Docs integration — runs in MAIN world (page context)
// Uses _docs_annotate_getAnnotatedText API for BOTH reading and writing text.
// This is the same approach as Lingdys — the annotate API is the authoritative
// document model, not the canvas rendering.

(function() {
  "use strict";

  let annotatedText = null;
  let lastEmittedText = "";

  // --- Annotate API initialization ---
  async function getAnnotatedText() {
    if (annotatedText) return annotatedText;
    if (typeof globalThis._docs_annotate_getAnnotatedText === "function") {
      try {
        annotatedText = await globalThis._docs_annotate_getAnnotatedText("nb");
        console.log("NorskTale: annotatedText API loaded");
      } catch (e) {
        console.log("NorskTale: annotatedText API failed: " + e);
      }
    }
    return annotatedText;
  }

  // Initialize the API as early as possible
  getAnnotatedText();

  // --- Text reading via annotate API ---
  async function emitText() {
    if (!document.hasFocus()) return;
    const iframe = document.querySelector("iframe.docs-texteventtarget-iframe");
    if (!iframe) return;

    const at = await getAnnotatedText();
    if (!at || typeof at.getText !== "function") return;

    let fullText;
    try {
      fullText = at.getText();
    } catch (e) {
      console.log("NorskTale: getText() error: " + e);
      return;
    }

    // Strip control characters (annotate API prepends \u0003 ETX)
    fullText = fullText.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");

    if (!fullText || fullText === lastEmittedText) return;
    lastEmittedText = fullText;

    // Get cursor position from selection
    let cursorIndex = fullText.length;
    try {
      const sel = at.getSelection();
      if (sel && sel.length > 0) {
        cursorIndex = sel[0].start || 0;
      }
    } catch (e) {}

    // Cross-world communication via DOM element
    let el = document.getElementById("norsktale-data");
    if (!el) {
      el = document.createElement("div");
      el.id = "norsktale-data";
      el.style.display = "none";
      document.documentElement.appendChild(el);
    }
    el.setAttribute("data-text", fullText);
    el.setAttribute("data-cursor", String(cursorIndex));
    el.dispatchEvent(new Event("norsktale-update", { bubbles: false }));
  }

  // --- Replace via annotate API ---
  function getIframeDoc() {
    const iframe = document.querySelector("iframe.docs-texteventtarget-iframe");
    if (!iframe) return null;
    try { return iframe.contentDocument; } catch(e) { return null; }
  }

  function insertReplacementText(text) {
    const doc = getIframeDoc();
    if (!doc) return false;
    const el = doc.querySelector("[contenteditable=true]");
    if (!el) return false;
    el.focus();
    const event = new InputEvent("beforeinput", {
      inputType: "insertReplacementText",
      data: text,
      isComposing: false
    });
    el.dispatchEvent(event);
    return true;
  }

  // Poll for replace requests from content.js (via DOM element)
  setInterval(() => {
    const replEl = document.getElementById("norsktale-replace");
    if (!replEl || replEl.getAttribute("data-pending") !== "true") return;
    replEl.setAttribute("data-pending", "false");
    const find = replEl.getAttribute("data-find");
    const replace = replEl.getAttribute("data-replace");
    const charOffset = parseInt(replEl.getAttribute("data-offset") || "0", 10);
    doReplace(find, replace, charOffset, replEl);
  }, 100);

  async function doReplace(find, replace, charOffset, replEl) {
    console.log("NorskTale gdocs-inject: replace '" + find + "' → '" + replace + "' at offset " + charOffset);

    const at = await getAnnotatedText();
    if (!at || typeof at.setSelection !== "function" || typeof at.getText !== "function") {
      console.log("NorskTale gdocs-inject: annotate API not available");
      if (replEl) { replEl.setAttribute("data-result", "false"); replEl.dispatchEvent(new Event("norsktale-replace-done")); }
      return;
    }

    try {
      const docText = at.getText();
      const findLower = find.toLowerCase();
      const docLower = docText.toLowerCase();

      // Find the occurrence closest to charOffset
      let bestPos = -1;
      let bestDist = Infinity;
      let searchFrom = 0;
      while (true) {
        const pos = docLower.indexOf(findLower, searchFrom);
        if (pos < 0) break;
        const dist = Math.abs(pos - charOffset);
        if (dist < bestDist) {
          bestDist = dist;
          bestPos = pos;
        }
        searchFrom = pos + 1;
      }

      if (bestPos >= 0) {
        console.log("NorskTale gdocs-inject: setSelection(" + bestPos + ", " + (bestPos + find.length) + ")");
        at.setSelection(bestPos, bestPos + find.length);
        const ok = insertReplacementText(replace);
        console.log("NorskTale gdocs-inject: insertReplacementText result: " + ok);
        // Clear lastEmittedText so next poll re-reads from annotate API
        lastEmittedText = "";
        if (replEl) { replEl.setAttribute("data-result", String(ok)); replEl.dispatchEvent(new Event("norsktale-replace-done")); }
        return;
      }
      console.log("NorskTale gdocs-inject: word not found in getText()");
    } catch (e) {
      console.log("NorskTale gdocs-inject: replace error: " + e);
    }

    if (replEl) { replEl.setAttribute("data-result", "false"); replEl.dispatchEvent(new Event("norsktale-replace-done")); }
  }

  // Periodic text emission — poll annotate API every 500ms
  setInterval(() => { emitText(); }, 500);

  console.log("NorskTale gdocs-inject.js loaded — annotate API text reading + replace active");
})();
