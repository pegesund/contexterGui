// NorskTale Google Docs integration — runs in MAIN world (page context)
// Uses _docs_annotate_getAnnotatedText API for BOTH reading and writing text.
// This is the same approach as Lingdys — the annotate API is the authoritative
// document model, not the canvas rendering.

(function() {
  "use strict";

  let annotatedText = null;
  let lastEmittedText = "";
  let lastEmittedCursor = -1;

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

    // Strip control characters (annotate API prepends \u0003 ETX) and trailing whitespace
    fullText = fullText.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "").trimEnd();

    if (!fullText) return;

    // Get cursor position — try annotate API first, fall back to DOM caret tracking
    let cursorIndex = fullText.length;
    try {
      const sel = at.getSelection();
      if (sel && sel.length > 0 && typeof sel[0].start === "number") {
        cursorIndex = sel[0].start;
      }
    } catch (e) {}

    // If annotate API returns end-of-text, try to find cursor from DOM caret element
    if (cursorIndex >= fullText.length - 1 || cursorIndex <= 0) {
      const domCursor = getDomCursorOffset(fullText);
      console.log("NorskTale: DOM cursor fallback: " + domCursor + " (API was " + cursorIndex + ", textLen=" + fullText.length + ")");
      if (domCursor >= 0) {
        cursorIndex = domCursor;
      }
    }

    // Skip if neither text nor cursor changed
    if (fullText === lastEmittedText && cursorIndex === lastEmittedCursor) return;
    lastEmittedText = fullText;
    lastEmittedCursor = cursorIndex;

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

  // --- DOM-based cursor position detection ---
  // Google Docs renders text in paragraphs (.kix-paragraphrenderer) with
  // line chunks (.kix-lineview). The user's caret is a div .kix-cursor-caret.
  // We find which paragraph the caret is in, get the text before the caret
  // within that paragraph, then search for that text in the annotate API text
  // to get the character offset.
  function getDomCursorOffset(fullText) {
    try {
      // Find the current user's caret
      const caret = document.querySelector(".kix-cursor-caret");
      if (!caret) return -1;
      const caretRect = caret.getBoundingClientRect();
      if (caretRect.height === 0) return -1;

      // Walk up to find the paragraph
      let para = caret.closest(".kix-paragraphrenderer");
      if (!para) return -1;

      // Get all text spans in this paragraph, in order
      const spans = para.querySelectorAll(".kix-wordhtmlgenerator-word-node");
      if (!spans.length) return -1;

      // Collect text before and at the caret position
      let textBefore = "";
      let caretX = caretRect.left;
      for (const span of spans) {
        const spanRect = span.getBoundingClientRect();
        // Span is entirely before caret
        if (spanRect.right <= caretX + 2) {
          textBefore += span.textContent;
        } else if (spanRect.left < caretX) {
          // Caret is within this span — estimate char position
          const spanText = span.textContent;
          const spanWidth = spanRect.width;
          if (spanWidth > 0 && spanText.length > 0) {
            const ratio = (caretX - spanRect.left) / spanWidth;
            const charIdx = Math.round(ratio * spanText.length);
            textBefore += spanText.substring(0, charIdx);
          }
          break;
        } else {
          break;
        }
      }

      // Now find this paragraph text in fullText to get the char offset
      // Get the full paragraph text
      let paraText = "";
      for (const span of spans) {
        paraText += span.textContent;
      }
      // Normalize whitespace for matching
      const paraClean = paraText.replace(/\u00a0/g, " ").trim();
      const fullClean = fullText.replace(/\u00a0/g, " ");

      if (paraClean.length === 0) return -1;

      // Find paragraph start in fullText
      const paraStart = fullClean.indexOf(paraClean);
      if (paraStart < 0) return -1;

      // Cursor offset = paragraph start + text before caret within paragraph
      const beforeClean = textBefore.replace(/\u00a0/g, " ");
      return paraStart + beforeClean.length;
    } catch (e) {
      return -1;
    }
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
