// NorskTale content script v8 — active cursor required
// Google Docs: text from canvas hook (gdocs-inject.js), replace via postMessage
// Other sites: textarea/contenteditable direct
// RULE: NEVER send text unless user has a blinking cursor in an editable element.

(function() {
  "use strict";
  console.log("NorskTale content.js v8 loaded");

  function norsktaleLog(msg) {
    console.log("NorskTale: " + msg);
    if (!port) connectPort();
    if (port) {
      try {
        port.postMessage({ type: "log", message: "[" + new Date().toLocaleTimeString() + "] " + msg });
      } catch(e) {}
    }
  }

  let lastSent = "";
  let activeElement = null;
  let lastTextElement = null;
  let lastGDocsText = null;
  let port = null;
  let replaceInProgress = false; // true while GDocs replace is pending
  let replaceStartTime = 0;

  function connectPort() {
    try {
      port = chrome.runtime.connect({ name: "norsktale-content" });
      port.onMessage.addListener(handleResponse);
      port.onDisconnect.addListener(() => { port = null; });
    } catch (e) { port = null; }
  }

  function isGoogleDocs() {
    return location.hostname === "docs.google.com" && location.pathname.startsWith("/document/");
  }

  // ========================================================================
  // ACTIVE CURSOR CHECK — the ONE gate for ALL text sending.
  // Returns the editable element with cursor, or null if no active cursor.
  // "Active cursor" means: user has clicked into an editable element and
  // there is a blinking caret / text selection inside it.
  // ========================================================================
  function getElementWithActiveCursor() {
    if (!document.hasFocus()) return null;

    const el = document.activeElement;
    if (!el) return null;

    // Textarea or text input: cursor exists if selectionStart is a number
    if (el.tagName === "TEXTAREA" || (el.tagName === "INPUT" && el.type === "text")) {
      if (typeof el.selectionStart === "number") return el;
      return null;
    }

    // ContentEditable: cursor exists if window.getSelection() has a range
    // AND that range is inside the editable element
    if (el.isContentEditable) {
      const sel = window.getSelection();
      if (!sel || sel.rangeCount === 0) return null;
      const range = sel.getRangeAt(0);
      if (el.contains(range.startContainer)) return el;
      return null;
    }

    // Not an editable element — no cursor
    return null;
  }

  // --- Google Docs: read text from DOM element written by gdocs-inject.js ---
  // gdocs-inject.js (MAIN world) writes to #norsktale-data.
  // Content.js (ISOLATED world) polls it. DOM is shared between worlds.
  let lastDataVersion = "";
  function pollGDocsData() {
    if (!isGoogleDocs()) return;
    // During replace: block until gdocs-inject confirms replacement (or timeout)
    if (replaceInProgress) {
      const el = document.getElementById("norsktale-data");
      const done = el && el.getAttribute("data-replace-done") === "true";
      const timedOut = Date.now() - replaceStartTime > 4000;
      if (done || timedOut) {
        norsktaleLog("Replace " + (done ? "confirmed" : "timed out") + " — resuming text updates");
        replaceInProgress = false;
        if (el) el.removeAttribute("data-replace-done");
      } else {
        return; // don't send stale text while replace is pending
      }
    }
    const el = document.getElementById("norsktale-data");
    if (!el) return;
    const text = el.getAttribute("data-text");
    const cursor = parseInt(el.getAttribute("data-cursor") || "0", 10);
    const caretX = parseInt(el.getAttribute("data-caret-x") || "0", 10);
    const caretY = parseInt(el.getAttribute("data-caret-y") || "0", 10);
    if (!text) return;
    const version = text + "|" + cursor + "|" + caretX + "|" + caretY;
    if (version === lastDataVersion) return;
    lastDataVersion = version;
    const data = { text: text, cursorStart: cursor, cursorEnd: cursor };
    lastGDocsText = data;
    const key = text + "|" + cursor;
    if (key === lastSent) return;
    lastSent = key;
    norsktaleLog("GDOCS canvas text: " + text.length + " chars, caret=(" + caretX + "," + caretY + ")");
    if (!port) connectPort();
    if (port) {
      try {
        port.postMessage({
          type: "textUpdate", text: text,
          cursorStart: cursor, cursorEnd: cursor,
          caretX: caretX, caretY: caretY, url: window.location.href
        });
      } catch (e) { port = null; }
    }
  }
  // Poll every 500ms for GDocs data
  setInterval(pollGDocsData, 500);

  // --- Replace handler ---
  function handleResponse(msg) {
    if (msg.action !== "replace") return;
    norsktaleLog("REPLACE: find='" + (msg.expected||"") + "' text='" + (msg.text||"") + "'");

    if (isGoogleDocs()) {
      // Send replace request to gdocs-inject.js (MAIN world) via postMessage
      const findText = msg.expected || "";
      const replaceText = msg.text || "";
      if (!findText) {
        norsktaleLog("REPLACE FAILED: no expected text");
        return;
      }
      // Send replace request via DOM element (shared between worlds)
      let replEl = document.getElementById("norsktale-replace");
      if (!replEl) {
        replEl = document.createElement("div");
        replEl.id = "norsktale-replace";
        replEl.style.display = "none";
        document.documentElement.appendChild(replEl);
      }
      replEl.setAttribute("data-find", findText);
      replEl.setAttribute("data-replace", replaceText);
      replEl.setAttribute("data-offset", String(msg.start || 0));
      replEl.setAttribute("data-pending", "true");
      replaceInProgress = true; // block text updates until replace confirmed
      replaceStartTime = Date.now();
      lastSent = ""; // force re-read after replace
      lastDataVersion = ""; // force re-read when fresh data arrives
      return;
    }

    // Non-GDocs: textarea or contenteditable
    const el = activeElement || lastTextElement;
    if (!el) return;
    let start = msg.start;
    let end = msg.end;
    const replacement = msg.text;
    const expected = msg.expected || "";

    if (el.tagName === "TEXTAREA" || el.tagName === "INPUT") {
      const val = el.value;
      if (expected && val.substring(start, end).toLowerCase() !== expected.toLowerCase()) {
        const idx = val.toLowerCase().indexOf(expected.toLowerCase(), Math.max(0, start - 5));
        if (idx >= 0 && idx <= start + 5) { start = idx; end = idx + expected.length; }
      }
      const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set
        || Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
      const newVal = val.substring(0, start) + replacement + val.substring(end);
      if (nativeSetter) nativeSetter.call(el, newVal);
      else el.value = newVal;
      el.selectionStart = start + replacement.length;
      el.selectionEnd = start + replacement.length;
      el.dispatchEvent(new Event("input", { bubbles: true }));
      el.dispatchEvent(new Event("change", { bubbles: true }));
      lastSent = "";
    } else if (el.isContentEditable) {
      el.focus();
      const sel = window.getSelection();
      const range = document.createRange();
      let startNode = null, startOfs = 0, endNode = null, endOfs = 0;
      let found = false;
      if (expected) {
        const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
        let buf = "";
        const nodes = [];
        while (walker.nextNode()) {
          nodes.push({ node: walker.currentNode, off: buf.length });
          buf += walker.currentNode.textContent;
        }
        const bufLower = buf.toLowerCase();
        const searchLower = expected.toLowerCase();
        let idx = bufLower.indexOf(searchLower, Math.max(0, start - 50));
        if (idx < 0) idx = bufLower.indexOf(searchLower);
        if (idx >= 0) {
          const endIdx = idx + expected.length;
          for (const n of nodes) {
            const nodeEnd = n.off + n.node.textContent.length;
            if (!startNode && idx < nodeEnd) { startNode = n.node; startOfs = idx - n.off; }
            if (endIdx <= nodeEnd) { endNode = n.node; endOfs = endIdx - n.off; break; }
          }
          if (startNode && endNode) found = true;
        }
      }
      if (found) {
        range.setStart(startNode, startOfs);
        range.setEnd(endNode, endOfs);
        sel.removeAllRanges();
        sel.addRange(range);
        document.execCommand("insertText", false, replacement) ||
          (range.deleteContents(),
           range.insertNode(document.createTextNode(replacement)),
           sel.collapseToEnd());
      }
      lastSent = "";
    }
  }

  // --- Non-GDocs: textarea/contenteditable monitoring ---
  function getTextAndCursor(el) {
    if (isGoogleDocs()) return null;
    if (!el) return null;
    if (el.tagName === "TEXTAREA" || (el.tagName === "INPUT" && el.type === "text")) {
      return { text: el.value, cursorStart: el.selectionStart, cursorEnd: el.selectionEnd, caretX: 0, caretY: 0 };
    }
    if (el.isContentEditable) {
      const sel = window.getSelection();
      const text = el.innerText;
      let cursorStart = 0, cursorEnd = 0;
      if (sel && sel.rangeCount > 0) {
        const range = sel.getRangeAt(0);
        const preRange = document.createRange();
        preRange.selectNodeContents(el);
        preRange.setEnd(range.startContainer, range.startOffset);
        cursorStart = preRange.toString().length;
        preRange.setEnd(range.endContainer, range.endOffset);
        cursorEnd = preRange.toString().length;
      }
      return { text, cursorStart, cursorEnd, caretX: 0, caretY: 0 };
    }
    return null;
  }

  // --- sendUpdate: ONLY sends if active cursor exists ---
  function sendUpdate() {
    if (isGoogleDocs()) return; // GDocs handled via postMessage from gdocs-inject.js
    const el = getElementWithActiveCursor();
    if (!el) return;
    const data = getTextAndCursor(el);
    if (!data || !data.text) return;
    const key = data.text + "|" + data.cursorStart;
    if (key === lastSent) return;
    lastSent = key;
    activeElement = el;
    lastTextElement = el;
    if (el.spellcheck !== false) el.spellcheck = false;
    if (!port) connectPort();
    if (port) {
      try {
        port.postMessage({
          type: "textUpdate", text: data.text,
          cursorStart: data.cursorStart, cursorEnd: data.cursorEnd,
          caretX: data.caretX, caretY: data.caretY, url: window.location.href
        });
      } catch (e) { port = null; }
    }
  }

  document.addEventListener("input", sendUpdate, true);
  document.addEventListener("selectionchange", sendUpdate, true);
  document.addEventListener("focusin", () => setTimeout(sendUpdate, 50), true);
  document.addEventListener("click", () => setTimeout(sendUpdate, 50), true);
  document.addEventListener("keyup", sendUpdate, true);

  // Clear stale text when page loses focus
  document.addEventListener("visibilitychange", () => {
    if (document.hidden) lastGDocsText = null;
  }, true);

  // Heartbeat — keeps native bridge alive, only sends text with active cursor
  setInterval(() => {
    if (!port) connectPort();
    if (isGoogleDocs()) {
      // GDocs: only resend if page has focus AND we have text from canvas hook
      // Don't resend stale text during replace
      if (replaceInProgress) {
        if (port) { try { port.postMessage({ type: "keepalive" }); } catch(e) { port = null; } }
        return;
      }
      if (document.hasFocus() && port && lastGDocsText) {
        try {
          port.postMessage({
            type: "textUpdate", text: lastGDocsText.text,
            cursorStart: lastGDocsText.cursorStart, cursorEnd: lastGDocsText.cursorEnd,
            caretX: 0, caretY: 0, url: window.location.href
          });
        } catch (e) { port = null; }
      } else if (port) {
        try { port.postMessage({ type: "keepalive" }); } catch(e) { port = null; }
      }
      return;
    }
    // Non-GDocs: ONLY send if active cursor exists
    const el = getElementWithActiveCursor();
    if (!el) return;
    const data = getTextAndCursor(el);
    if (!data || !data.text) return;
    if (port) {
      try {
        port.postMessage({
          type: "textUpdate", text: data.text,
          cursorStart: data.cursorStart, cursorEnd: data.cursorEnd,
          caretX: data.caretX, caretY: data.caretY, url: window.location.href
        });
      } catch (e) { port = null; }
    }
  }, 2000);
})();
