// NorskTale Google Docs integration — runs in MAIN world (page context)
// Hooks canvas fillText to capture text. Handles replace via insertReplacementText.
// No sidebar, no Apps Script.

(function() {
  "use strict";

  // Captured text chunks with positions
  let textChunks = [];  // { text, x, y, font, canvasId, canvas }
  let debounceTimer = null;
  let lastEmittedText = "";
  let canvasCounter = 0;
  const canvasIds = new WeakMap();
  const canvasRefs = new Map(); // canvasId -> canvas element

  function getCanvasId(canvas) {
    if (!canvasIds.has(canvas)) {
      canvasIds.set(canvas, ++canvasCounter);
      canvasRefs.set(canvasCounter, canvas);
    }
    return canvasIds.get(canvas);
  }

  // --- Canvas hooks ---

  const origFillText = CanvasRenderingContext2D.prototype.fillText;
  CanvasRenderingContext2D.prototype.fillText = function(text, x, y) {
    if (this.canvas && this.canvas.classList &&
        this.canvas.classList.contains("kix-canvas-tile-content")) {
      const transform = this.getTransform();
      const absX = transform.e + x * transform.a;
      const absY = transform.f + y * transform.d;
      const cid = getCanvasId(this.canvas);

      textChunks.push({
        text: text, x: absX, y: absY,
        font: this.font || "",
        canvasId: cid, canvas: this.canvas
      });

      if (debounceTimer) clearTimeout(debounceTimer);
      debounceTimer = setTimeout(emitText, 150);
    }
    return origFillText.apply(this, arguments);
  };

  const origClearRect = CanvasRenderingContext2D.prototype.clearRect;
  CanvasRenderingContext2D.prototype.clearRect = function(x, y, w, h) {
    if (this.canvas && this.canvas.classList &&
        this.canvas.classList.contains("kix-canvas-tile-content")) {
      const cw = this.canvas.width;
      const ch = this.canvas.height;
      if (x <= 0 && y <= 0 && w >= cw && h >= ch) {
        const cid = getCanvasId(this.canvas);
        textChunks = textChunks.filter(c => c.canvasId !== cid);
      }
    }
    return origClearRect.apply(this, arguments);
  };

  // --- Text extraction ---

  function stripBidi(text) {
    return text.replace(/[\u200E\u200F\u202A-\u202E\u2066-\u2069\u200B-\u200D\uFEFF\u2060]/g, "");
  }

  // Deduplicate chunks: overlapping canvas layers produce identical text at same (x,y).
  // Keep only the first chunk at each (text, roundedX, roundedY).
  function deduplicateChunks(chunks) {
    const seen = new Set();
    const result = [];
    for (const c of chunks) {
      const key = c.text + "|" + Math.round(c.x) + "|" + Math.round(c.y);
      if (!seen.has(key)) {
        seen.add(key);
        result.push(c);
      }
    }
    return result;
  }

  // Build structured lines from chunks
  function buildLines() {
    if (textChunks.length === 0) return [];
    const unique = deduplicateChunks(textChunks);
    const sorted = [...unique].sort((a, b) => a.y - b.y || a.x - b.x);
    const lines = [];
    let currentLine = [sorted[0]];
    let currentY = sorted[0].y;

    for (let i = 1; i < sorted.length; i++) {
      if (Math.abs(sorted[i].y - currentY) < 3) {
        currentLine.push(sorted[i]);
      } else {
        currentLine.sort((a, b) => a.x - b.x);
        lines.push(currentLine);
        currentLine = [sorted[i]];
        currentY = sorted[i].y;
      }
    }
    currentLine.sort((a, b) => a.x - b.x);
    lines.push(currentLine);
    return lines;
  }

  // Deduplicate text that repeats due to overlapping canvas layers.
  // Google Docs renders on multiple canvas layers — we capture all of them.
  // Detect if text is N copies of itself and return just one copy.
  function deduplicateText(text) {
    if (text.length < 10) return text;
    const len = text.length;
    // Try dividing by 2, 3, 4
    for (const n of [2, 3, 4]) {
      if (len % n !== 0) continue;
      const partLen = len / n;
      const part = text.substring(0, partLen);
      let match = true;
      for (let i = 1; i < n; i++) {
        if (text.substring(i * partLen, (i + 1) * partLen) !== part) {
          match = false;
          break;
        }
      }
      if (match) return part;
    }
    // Try line-based dedup: if lines repeat in blocks
    const lines = text.split("\n");
    for (const n of [2, 3, 4]) {
      if (lines.length % n !== 0) continue;
      const blockSize = lines.length / n;
      if (blockSize < 1) continue;
      const block = lines.slice(0, blockSize).join("\n");
      let match = true;
      for (let i = 1; i < n; i++) {
        if (lines.slice(i * blockSize, (i + 1) * blockSize).join("\n") !== block) {
          match = false;
          break;
        }
      }
      if (match) return block;
    }
    return text;
  }

  function emitText() {
    const lines = buildLines();
    if (lines.length === 0) return;

    let fullText = lines.map(line =>
      line.map(c => stripBidi(c.text)).join("")
    ).join("\n");

    fullText = deduplicateText(fullText);

    // ONLY emit if user has an active cursor in the document.
    // Check 1: document must have focus
    if (!document.hasFocus()) return;
    // Check 2: the docs editing iframe must exist and be active
    const iframe = document.querySelector("iframe.docs-texteventtarget-iframe");
    if (!iframe) return;

    if (fullText === lastEmittedText) return;
    lastEmittedText = fullText;

    // Try to find cursor position from caret element (best effort)
    let cursorIndex = fullText.length;
    const caret = document.querySelector(".kix-cursor-caret, #kix-current-user-cursor-caret, [class*='kix-cursor']");
    if (caret) {
      const caretRect = caret.getBoundingClientRect();
      if (caretRect.height > 0) {
        cursorIndex = findCharIndexAtPoint(caretRect.left, caretRect.top + caretRect.height / 2, lines);
      }
    }

    // Cross-world communication via DOM element (postMessage and CustomEvent.detail
    // do NOT cross MAIN→ISOLATED world boundary in MV3)
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

  // Find character index in the full text at a given viewport point
  function findCharIndexAtPoint(px, py, lines) {
    let charIndex = 0;
    for (const line of lines) {
      const lineText = line.map(c => stripBidi(c.text)).join("");
      // Check if this line is at the right Y
      const chunk = line[0];
      const rect = chunk.canvas.getBoundingClientRect();
      const scale = rect.height / chunk.canvas.height;
      const chunkPageY = rect.top + chunk.y * scale;

      if (Math.abs(chunkPageY - py) < 20) {
        // This line — find X position
        let lineCharPos = 0;
        for (const c of line) {
          const cRect = c.canvas.getBoundingClientRect();
          const cScale = cRect.width / c.canvas.width;
          const chunkPageX = cRect.left + c.x * cScale;
          if (chunkPageX > px) break;
          lineCharPos += stripBidi(c.text).length;
        }
        return charIndex + Math.min(lineCharPos, lineText.length);
      }
      charIndex += lineText.length + 1;
    }
    return charIndex;
  }

  // --- Replace: Lingdys-style approach ---
  // 1. Click at word position to place cursor
  // 2. Use Ctrl+Shift+Left to select word backward (or Shift+Right to select forward)
  // 3. Use execCommand('insertText') on the iframe to type replacement

  // Save original measureText before any hooks
  const origMeasureText = CanvasRenderingContext2D.prototype.measureText;

  function getIframeDoc() {
    const iframe = document.querySelector("iframe.docs-texteventtarget-iframe");
    if (!iframe) return null;
    try { return iframe.contentDocument; } catch(e) { return null; }
  }

  function getEditor() {
    return document.querySelector(".kix-appview-editor");
  }

  // Find the viewport coordinates of a word in the captured text
  function findWordCoords(word, charOffset) {
    const lines = buildLines();
    let globalCharIdx = 0;

    for (const line of lines) {
      const lineText = line.map(c => stripBidi(c.text)).join("");
      const lineStart = globalCharIdx;

      const wordLower = word.toLowerCase();
      const lineLower = lineText.toLowerCase();
      let pos = lineLower.indexOf(wordLower);

      while (pos >= 0) {
        const wordGlobalStart = lineStart + pos;
        if (charOffset === undefined || Math.abs(wordGlobalStart - charOffset) < 200) {
          const startCoords = getChunkCoordsAtLinePos(line, pos);
          const endCoords = getChunkCoordsAtLinePos(line, pos + word.length);
          if (startCoords && endCoords) {
            return { startX: startCoords.x, startY: startCoords.y,
                     endX: endCoords.x, endY: endCoords.y };
          }
        }
        pos = lineLower.indexOf(wordLower, pos + 1);
      }
      globalCharIdx += lineText.length + 1;
    }
    return null;
  }

  function getChunkCoordsAtLinePos(line, charPos) {
    let pos = 0;
    for (const chunk of line) {
      const cleanText = stripBidi(chunk.text);
      if (pos + cleanText.length >= charPos) {
        const localPos = charPos - pos;
        const fraction = cleanText.length > 0 ? localPos / cleanText.length : 0;
        const rect = chunk.canvas.getBoundingClientRect();
        const scaleX = rect.width / chunk.canvas.width;
        const scaleY = rect.height / chunk.canvas.height;
        const ctx = chunk.canvas.getContext("2d");
        let chunkWidth = 10;
        if (ctx) {
          ctx.font = chunk.font;
          chunkWidth = origMeasureText.call(ctx, cleanText).width * scaleX;
        }
        return {
          x: rect.left + chunk.x * scaleX + fraction * chunkWidth,
          y: rect.top + chunk.y * scaleY
        };
      }
      pos += cleanText.length;
    }
    if (line.length > 0) {
      const last = line[line.length - 1];
      const rect = last.canvas.getBoundingClientRect();
      return { x: rect.left + last.x * (rect.width / last.canvas.width) + 10,
               y: rect.top + last.y * (rect.height / last.canvas.height) };
    }
    return null;
  }

  // Click at a position in the editor
  function clickAt(x, y, opts_extra) {
    const editor = getEditor();
    if (!editor) return;
    const opts = Object.assign({ bubbles: true, cancelable: true, button: 0,
                   clientX: x, clientY: y, composed: true }, opts_extra || {});
    editor.dispatchEvent(new MouseEvent("mousedown", opts));
    editor.dispatchEvent(new MouseEvent("mouseup", opts));
  }

  // Double-click at a position (selects word in Google Docs)
  function doubleClickAt(x, y) {
    const editor = getEditor();
    if (!editor) return;
    const opts = { bubbles: true, cancelable: true, button: 0,
                   clientX: x, clientY: y, detail: 2, composed: true };
    editor.dispatchEvent(new MouseEvent("mousedown", opts));
    editor.dispatchEvent(new MouseEvent("mouseup", opts));
    editor.dispatchEvent(new MouseEvent("dblclick", opts));
  }

  // Type text into Google Docs by simulating a paste event.
  // Google Docs has its own paste handler that reads clipboardData —
  // this works even with synthetic (untrusted) events.
  function typeText(text) {
    const doc = getIframeDoc();
    if (!doc) return false;
    const el = doc.querySelector("[contenteditable=true]");
    if (!el) return false;
    el.focus();

    const dt = new DataTransfer();
    dt.setData("text/plain", text);
    const pasteEvent = new ClipboardEvent("paste", {
      clipboardData: dt,
      bubbles: true,
      cancelable: true
    });
    el.dispatchEvent(pasteEvent);
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

  function doReplace(find, replace, charOffset, replEl) {
    console.log("NorskTale gdocs-inject: replace '" + find + "' → '" + replace + "'");

    const coords = findWordCoords(find, charOffset);
    if (!coords) {
      console.log("NorskTale gdocs-inject: word not found in canvas text");
      if (replEl) { replEl.setAttribute("data-result", "false"); replEl.dispatchEvent(new Event("norsktale-replace-done")); }
      return;
    }

    // Click at start, shift-click at end (pull endX back 2px to avoid trailing space)
    console.log("NorskTale gdocs-inject: select " + coords.startX.toFixed(0) + "," +
                coords.startY.toFixed(0) + " → " + coords.endX.toFixed(0) + "," + coords.endY.toFixed(0));
    clickAt(coords.startX, coords.startY);

    setTimeout(() => {
      clickAt(coords.endX - 2, coords.endY, { shiftKey: true });

      setTimeout(() => {
        const ok = typeText(replace);
        console.log("NorskTale gdocs-inject: typeText result: " + ok);
        if (replEl) { replEl.setAttribute("data-result", String(ok)); replEl.dispatchEvent(new Event("norsktale-replace-done")); }
      }, 150);
    }, 150);
  }

  // Periodic re-emit
  setInterval(() => {
    if (textChunks.length > 0) emitText();
  }, 2000);

  console.log("NorskTale gdocs-inject.js loaded — canvas text capture + direct replace active");
})();
