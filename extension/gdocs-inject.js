// Spell Google Docs integration — runs in MAIN world (page context)
// Uses _docs_annotate_getAnnotatedText API for BOTH reading and writing text.
// This is the same approach as Lingdys — the annotate API is the authoritative
// document model, not the canvas rendering.

(function() {
  "use strict";

  let lastEmittedText = "";
  let lastEmittedCursor = -1;
  let lastParaStart = 0; // document offset where the last emitted paragraph begins

  // --- Annotate API ---
  //
  // _docs_annotate_getAnnotatedText("nb") returns an object whose
  // getText() / getSelection() methods read from a snapshot Google takes
  // at factory-call time. Google refreshes that snapshot on cursor /
  // click events but NOT on pure keystrokes inside the texteventtarget
  // iframe — so any factory result we cache silently stales out the
  // moment the user starts typing.
  //
  // The previous fix (2026-05-19) tried to attach a keydown listener to
  // the GDocs input iframe and null the cache on every keystroke.  User
  // reported it didn't help — likely because that iframe's contentDocument
  // is locked down in some Google Docs surfaces (sandboxed / cross-origin
  // wrapper / replaced before our listener fires).
  //
  // Simpler robust fix: don't cache at all.  Call the factory every time
  // emitText() runs (~ once per 500 ms).  The factory call is cheap —
  // Google's own UI invokes it constantly — and it always returns a
  // fresh snapshot, so:
  //   - typing → next 500 ms tick sees the new text
  //   - tab-switch and return → next tick sees current state
  //   - Cmd+A + Backspace → next tick sees empty text
  // No more "click between words to refresh" workaround needed.
  async function getAnnotatedText() {
    if (typeof globalThis._docs_annotate_getAnnotatedText !== "function") {
      return null;
    }
    try {
      return await globalThis._docs_annotate_getAnnotatedText("nb");
    } catch (e) {
      console.log("Spell: annotatedText API failed: " + e);
      return null;
    }
  }

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

  // --- DOM-based primary text source ---
  //
  // Google Docs renders to a CANVAS, not to DOM elements, but it still
  // maintains a DOM mirror (.kix-paragraphrenderer with
  // .kix-wordhtmlgenerator-word-node spans) for accessibility / screen
  // readers, AND that mirror is updated synchronously on every keystroke.
  // The _docs_annotate_getAnnotatedText("nb") factory, by contrast,
  // returns a snapshot of Google's internal "annotated text" model
  // which Google refreshes lazily — it can stay STALE through
  // Cmd+A+Backspace + retype until the user clicks somewhere, which is
  // exactly the user-reported 2026-05-19 bug:
  //   "Then I type something and even press spacebar nothing happens.
  //    App keeps showing previous errors until I press on a word using
  //    my cursor which trigger's something to refresh our app."
  //
  // Read text from the DOM mirror primarily; use the annotate API only
  // as a fallback for surfaces where no .kix-paragraphrenderer is
  // present (read-only viewers, hydration not finished).
  function getDomFullText() {
    try {
      const paragraphs = document.querySelectorAll(".kix-paragraphrenderer");
      if (paragraphs.length === 0) return null;
      const parts = [];
      for (const para of paragraphs) {
        const spans = para.querySelectorAll(".kix-wordhtmlgenerator-word-node");
        let paraText = "";
        for (const span of spans) {
          paraText += span.textContent;
        }
        // Normalize non-breaking spaces to ASCII so downstream
        // word-splitting works the same as for textareas.
        parts.push(paraText.replace(/ /g, " "));
      }
      return parts.join("\n");
    } catch (e) {
      return null;
    }
  }

  async function emitText() {
    if (!document.hasFocus()) return;
    const iframe = document.querySelector("iframe.docs-texteventtarget-iframe");
    if (!iframe) return;

    let fullText = getDomFullText();
    // Fallback: only ask the annotate API if the DOM mirror produced no
    // paragraphs at all.  That keeps us off the stale snapshot Google
    // returns from _docs_annotate_getAnnotatedText after Cmd+A+Backspace.
    if (fullText === null) {
      const at = await getAnnotatedText();
      if (!at || typeof at.getText !== "function") return;
      try {
        fullText = at.getText();
      } catch (e) {
        console.log("Spell: getText() error: " + e);
        return;
      }
    }

    // Strip control characters (annotate API prepends \u0003 ETX)
    // Do NOT trimEnd — trailing space after a word is needed to detect word completion
    fullText = fullText.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, "");

    // Previously we bailed here when fullText was empty
    // ("if (!fullText) return;") — same family of bug as content.js:
    // after Cmd+A + Backspace inside a Google Doc the document went
    // empty, we suppressed the update, and the desktop kept showing
    // the stale errors. Reported 2026-05-19: "this issue still
    // exists in docs". Allow empty text through; the desktop's
    // prune_resolved_errors empty-doc branch (main.rs:3335) handles
    // it. We still dedup on (paraText, cursorInPara) below so a
    // steady-state empty doc only emits once.

    // Cursor position: DOM is the primary source.  The annotate API's
    // getSelection() has the same stale-snapshot problem as its getText()
    // — after Cmd+A+Backspace it can return the OLD cursor position
    // until the user clicks. getDomCursorOffset finds the live
    // .kix-cursor-caret DOM element + works out the character offset
    // by walking spans, which always reflects the current state.
    let cursorIndex = getDomCursorOffset(fullText);
    if (cursorIndex < 0) {
      // No live caret in DOM (e.g. window unfocused mid-poll) — try
      // the annotate API selection as a fallback.
      cursorIndex = fullText.length;
      try {
        const at = await getAnnotatedText();
        if (at && typeof at.getSelection === "function") {
          const sel = at.getSelection();
          if (sel && sel.length > 0 && typeof sel[0].start === "number") {
            cursorIndex = sel[0].start;
          }
        }
      } catch (e) {}
    }

    // Extract only the paragraph the cursor is in (iOS-style: send only the active block)
    const { paraText, cursorInPara, paraStart } = extractParagraphAtCursor(fullText, cursorIndex);
    lastParaStart = paraStart; // remember for doReplace offset conversion

    // Skip if neither paragraph text nor cursor-within-paragraph changed
    if (paraText === lastEmittedText && cursorInPara === lastEmittedCursor) return;
    lastEmittedText = paraText;
    lastEmittedCursor = cursorInPara;

    // Cross-world communication via DOM element
    let el = document.getElementById("spell-data");
    if (!el) {
      el = document.createElement("div");
      el.id = "spell-data";
      el.style.display = "none";
      document.documentElement.appendChild(el);
    }
    // Get caret screen position for window-follows-cursor.
    //
    // IMPORTANT: emit PHYSICAL pixels (multiply logical coords by
    // devicePixelRatio), matching what content.js's viewportToScreen() does
    // for textareas / contenteditable. Without the `* dpr` multiplier this
    // path produced LOGICAL points while the textarea path produced
    // physical, so the desktop side (which now divides bridge caret by
    // pixels_per_point on macOS to recover logical) over-corrected for
    // Google Docs and the Spell window jumped to the top of the screen —
    // observed 2026-05-19 in SS 3 ("Hello I am doi" doc with the window
    // floating at the top of the page instead of below the caret line).
    let caretScreenX = 0, caretScreenY = 0;
    try {
      const dpr = window.devicePixelRatio || 1;
      const chromeHeight = window.outerHeight - window.innerHeight;
      const toPhysical = (lx, ly) => {
        caretScreenX = Math.round((window.screenX + lx) * dpr);
        caretScreenY = Math.round((window.screenY + chromeHeight + ly + 5) * dpr);
      };
      // Primary: kix-cursor-caret (the blinking I-beam)
      const caret = document.querySelector(".kix-cursor-caret");
      if (caret) {
        const r = caret.getBoundingClientRect();
        if (r.height > 0) toPhysical(r.left, r.bottom);
      }
      // Fallback 1: kix-cursor element (parent of kix-cursor-caret)
      if (!caretScreenX && !caretScreenY) {
        const cursor = document.querySelector(".kix-cursor");
        if (cursor) {
          const r = cursor.getBoundingClientRect();
          if (r.height > 0) toPhysical(r.left, r.bottom);
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
              if (r && r.height > 0) toPhysical(r.left, r.bottom);
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
    el.dispatchEvent(new Event("spell-update", { bubbles: false }));
  }

  // --- DOM-based cursor position detection ---
  // Google Docs renders text in paragraphs (.kix-paragraphrenderer) with
  // line chunks (.kix-lineview). The user's caret is a div .kix-cursor-caret.
  // We find which paragraph the caret is currently OVER (by spatial
  // bounding-rect lookup), get the text before the caret within that
  // paragraph, then locate that text in fullText to compute the
  // document-level character offset.
  //
  // IMPORTANT: do NOT use `caret.closest(".kix-paragraphrenderer")` — the
  // .kix-cursor-caret element's DOM PARENT only updates when the user
  // clicks somewhere; pure keystrokes move the caret by updating its CSS
  // transform (which getBoundingClientRect reflects) but DO NOT reparent
  // it.  So .closest() would return the OLD paragraph the user clicked
  // into, while getBoundingClientRect already shows the visual position
  // at the current typing location.  Reported 2026-05-19: "bulb mode
  // does not seems to be properly working with google docs. It does not
  // show any suggestions while I am writing. I have to click on the
  // text then it show suggestions for the current word I am typing."
  function getDomCursorOffset(fullText) {
    try {
      // Find the current user's caret
      const caret = document.querySelector(".kix-cursor-caret");
      if (!caret) return -1;
      const caretRect = caret.getBoundingClientRect();
      if (caretRect.height === 0) return -1;

      // Spatial paragraph lookup: pick the .kix-paragraphrenderer whose
      // bounding rect actually contains the caret's screen position.
      // Use a midpoint Y inside the caret so a caret sitting exactly on
      // a paragraph boundary picks the line it's drawing on.
      const caretMidY = caretRect.top + caretRect.height / 2;
      const caretX = caretRect.left;
      let para = null;
      const allParagraphs = document.querySelectorAll(".kix-paragraphrenderer");
      for (const p of allParagraphs) {
        const r = p.getBoundingClientRect();
        if (caretMidY >= r.top && caretMidY <= r.bottom
            && caretX >= r.left - 2 && caretX <= r.right + 2) {
          para = p;
          break;
        }
      }
      // Last-resort fallback to the old behaviour — if no spatial hit
      // (e.g. caret floating outside any rendered paragraph during a
      // transition), keep going with the DOM-parent guess so we at
      // least produce A position rather than bailing entirely.
      if (!para) para = caret.closest(".kix-paragraphrenderer");
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
    const replEl = document.getElementById("spell-replace");
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
    console.log("Spell gdocs-inject: replace '" + find + "' → '" + replace + "' paraOffset=" + charOffset + " docOffset=" + docAbsOffset);

    const at = await getAnnotatedText();
    if (!at || typeof at.setSelection !== "function" || typeof at.getText !== "function") {
      console.log("Spell gdocs-inject: annotate API not available");
      if (replEl) { replEl.setAttribute("data-result", "false"); replEl.dispatchEvent(new Event("spell-replace-done")); }
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
        console.log("Spell gdocs-inject: setSelection(" + bestPos + ", " + (bestPos + find.length) + ")");
        at.setSelection(bestPos, bestPos + find.length);
        const ok = insertReplacementText(replace);
        console.log("Spell gdocs-inject: insertReplacementText result: " + ok);
        // Clear lastEmittedText so next poll re-reads from annotate API
        lastEmittedText = "";
        if (replEl) { replEl.setAttribute("data-result", String(ok)); replEl.dispatchEvent(new Event("spell-replace-done")); }
        return;
      }
      console.log("Spell gdocs-inject: word not found in getText()");
    } catch (e) {
      console.log("Spell gdocs-inject: replace error: " + e);
    }

    if (replEl) { replEl.setAttribute("data-result", "false"); replEl.dispatchEvent(new Event("spell-replace-done")); }
  }

  // Periodic text emission — poll annotate API every 500ms
  setInterval(() => { emitText(); }, 500);

  console.log("Spell gdocs-inject.js loaded — annotate API text reading + replace active");
})();
