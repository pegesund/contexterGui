// NorskTale Google Docs integration — runs in MAIN world (page context)
// Uses _docs_annotate_getAnnotatedText API for BOTH reading and writing text.
// This is the same approach as Lingdys — the annotate API is the authoritative
// document model, not the canvas rendering.

(function() {
  "use strict";

  let annotatedText = null;
  let lastEmittedText = "";
  let lastEmittedCursor = -1;
  let lastParaStart = 0; // document offset where the last emitted paragraph begins

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

  // --- Extract the paragraph containing the cursor ---
  // Paragraphs in Google Docs are separated by \n.
  // Returns { paraText, cursorInPara } so the Rust app only sees the active paragraph,
  // mirroring the iOS approach of sending only the relevant text block.
  function extractParagraphAtCursor(fullText, cursorIndex) {
    // Clamp cursor to valid range
    const cursor = Math.max(0, Math.min(cursorIndex, fullText.length));

    // Find paragraph start: last \n before cursor
    const paraStart = fullText.lastIndexOf("\n", cursor - 1) + 1;

    // Find paragraph end: next \n after cursor (or end of text)
    const nextNewline = fullText.indexOf("\n", cursor);
    const paraEnd = nextNewline === -1 ? fullText.length : nextNewline;

    const paraText = fullText.substring(paraStart, paraEnd);
    const cursorInPara = cursor - paraStart;

    return { paraText, cursorInPara, paraStart };
  }

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
    // Do NOT trimEnd — trailing space after a word is needed to detect word completion
    fullText = fullText.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");

    if (!fullText) return;

    // Get cursor position in full document — try annotate API first, fall back to DOM
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

    // Extract only the paragraph the cursor is in (iOS-style: send only the active block)
    const { paraText, cursorInPara, paraStart } = extractParagraphAtCursor(fullText, cursorIndex);
    lastParaStart = paraStart; // remember for doReplace offset conversion

    // Skip if neither paragraph text nor cursor-within-paragraph changed
    if (paraText === lastEmittedText && cursorInPara === lastEmittedCursor) return;
    lastEmittedText = paraText;
    lastEmittedCursor = cursorInPara;

    // Cross-world communication via DOM element
    let el = document.getElementById("norsktale-data");
    if (!el) {
      el = document.createElement("div");
      el.id = "norsktale-data";
      el.style.display = "none";
      document.documentElement.appendChild(el);
    }
    // Get caret screen position for window-follows-cursor
    let caretScreenX = 0, caretScreenY = 0;
    try {
      const chromeHeight = window.outerHeight - window.innerHeight;
      // Primary: kix-cursor-caret (the blinking I-beam)
      const caret = document.querySelector(".kix-cursor-caret");
      if (caret) {
        const r = caret.getBoundingClientRect();
        if (r.height > 0) {
          caretScreenX = Math.round(window.screenX + r.left);
          caretScreenY = Math.round(window.screenY + chromeHeight + r.bottom + 5);
        }
      }
      // Fallback 1: kix-cursor element (parent of kix-cursor-caret)
      if (!caretScreenX && !caretScreenY) {
        const cursor = document.querySelector(".kix-cursor");
        if (cursor) {
          const r = cursor.getBoundingClientRect();
          if (r.height > 0) {
            caretScreenX = Math.round(window.screenX + r.left);
            caretScreenY = Math.round(window.screenY + chromeHeight + r.bottom + 5);
          }
        }
      }
      // Fallback 2: use the selection range from the iframe
      if (!caretScreenX && !caretScreenY) {
        const iframe = document.querySelector("iframe.docs-texteventtarget-iframe");
        if (iframe) {
          try {
            const iDoc = iframe.contentDocument;
            const iSel = iDoc && iDoc.getSelection && iDoc.getSelection();
            if (iSel && iSel.rangeCount > 0) {
              const r = iSel.getRangeAt(0).getBoundingClientRect();
              if (r && r.height > 0) {
                caretScreenX = Math.round(window.screenX + r.left);
                caretScreenY = Math.round(window.screenY + chromeHeight + r.bottom + 5);
              }
            }
          } catch(e2) {}
        }
      }
    } catch (e) {}

    // Send only the active paragraph + paragraph-relative cursor (iOS approach)
    el.setAttribute("data-text", paraText);
    el.setAttribute("data-cursor", String(cursorInPara));
    el.setAttribute("data-caret-x", String(caretScreenX));
    el.setAttribute("data-caret-y", String(caretScreenY));
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
    // charOffset is paragraph-relative — convert to document-absolute for the annotate API
    const docAbsOffset = lastParaStart + charOffset;
    console.log("NorskTale gdocs-inject: replace '" + find + "' → '" + replace + "' paraOffset=" + charOffset + " docOffset=" + docAbsOffset);

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

      // Find the occurrence closest to the document-absolute offset
      let bestPos = -1;
      let bestDist = Infinity;
      let searchFrom = 0;
      while (true) {
        const pos = docLower.indexOf(findLower, searchFrom);
        if (pos < 0) break;
        const dist = Math.abs(pos - docAbsOffset);
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
