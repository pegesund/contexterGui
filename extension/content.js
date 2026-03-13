// NorskTale content script v7 — no sidebar, no Apps Script
// Google Docs: text from canvas hook (gdocs-inject.js), replace via insertReplacementText
// Other sites: textarea/contenteditable direct

(function() {
  "use strict";
  console.log("NorskTale content.js v7 loaded");

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

  // --- Google Docs: receive text from canvas hook ---
  document.addEventListener("norsktale-gdocs-text", function(e) {
    const data = e.detail;
    if (!data || !data.text) return;
    lastGDocsText = data;
    const key = data.text + "|" + data.cursorStart;
    if (key === lastSent) return;
    lastSent = key;
    norsktaleLog("GDOCS canvas text: " + data.text.length + " chars");
    if (!port) connectPort();
    if (port) {
      try {
        port.postMessage({
          type: "textUpdate", text: data.text,
          cursorStart: data.cursorStart, cursorEnd: data.cursorEnd,
          caretX: 0, caretY: 0, url: window.location.href
        });
      } catch (e) { port = null; }
    }
  });

  // --- Replace handler ---
  function handleResponse(msg) {
    if (msg.action !== "replace") return;
    norsktaleLog("REPLACE: find='" + (msg.expected||"") + "' text='" + (msg.text||"") + "'");

    if (isGoogleDocs()) {
      // Send replace request to gdocs-inject.js (MAIN world) via CustomEvent
      const findText = msg.expected || "";
      const replaceText = msg.text || "";
      if (!findText) {
        norsktaleLog("REPLACE FAILED: no expected text");
        return;
      }
      document.dispatchEvent(new CustomEvent("norsktale-gdocs-replace", {
        detail: { find: findText, replace: replaceText, charOffset: msg.start || 0 }
      }));
      // Listen for result
      document.addEventListener("norsktale-gdocs-replace-result", function handler(e) {
        document.removeEventListener("norsktale-gdocs-replace-result", handler);
        norsktaleLog("REPLACE result: " + JSON.stringify(e.detail));
        lastSent = ""; // force re-read after replace
      });
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

  function sendUpdate() {
    if (!document.hasFocus()) return;
    if (isGoogleDocs()) return;
    const el = document.activeElement;
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

  // Heartbeat
  setInterval(() => {
    if (!document.hasFocus()) return;
    if (isGoogleDocs() && lastGDocsText) {
      if (!port) connectPort();
      if (port) {
        try {
          port.postMessage({
            type: "textUpdate", text: lastGDocsText.text,
            cursorStart: lastGDocsText.cursorStart, cursorEnd: lastGDocsText.cursorEnd,
            caretX: 0, caretY: 0, url: window.location.href
          });
        } catch (e) { port = null; }
      }
      return;
    }
    const el = activeElement || lastTextElement || document.activeElement;
    if (!el) return;
    const data = getTextAndCursor(el);
    if (!data || !data.text) return;
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
  }, 2000);
})();
