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
  let replaceInProgress = false;
  let replaceStartTime = 0;
  let lastParaStart = 0;

  function extractParagraphAtCursor(text, cursorPos) {
    const cursor = Math.max(0, Math.min(cursorPos, text.length));
    const paraStart = text.lastIndexOf("\n", cursor - 1) + 1;
    const nextNewline = text.indexOf("\n", cursor);
    const paraEnd = nextNewline === -1 ? text.length : nextNewline;
    return {
      paraText: text.substring(paraStart, paraEnd),
      cursorInPara: cursor - paraStart,
      paraStart
    };
  }

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

  function isGmail() {
    return location.hostname === "mail.google.com";
  }

  // Walk up the DOM to find the nearest contenteditable ancestor.
  // Needed for Gmail compose where document.activeElement is a wrapper div,
  // not the actual textbox the cursor is in.
  function findContentEditableAncestor(node) {
    while (node && node !== document.body) {
      if (node.nodeType === 1 && node.isContentEditable) return node;
      node = node.parentNode;
    }
    return null;
  }

  function getElementWithActiveCursor() {
    if (!document.hasFocus()) return null;

    const el = document.activeElement;

    if (el && (el.tagName === "TEXTAREA" || (el.tagName === "INPUT" && el.type === "text"))) {
      if (typeof el.selectionStart === "number") return el;
    }

    if (el && el.isContentEditable) {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0) {
        const range = sel.getRangeAt(0);
        if (el.contains(range.startContainer)) return el;
      }
    }

    // Fallback for Gmail and rich editors where activeElement ≠ the textbox
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      const anchor = sel.getRangeAt(0).startContainer;
      const editable = findContentEditableAncestor(
        anchor.nodeType === 3 ? anchor.parentNode : anchor
      );
      if (editable) return editable;
    }

    return null;
  }

  // gdocs-inject.js (MAIN world) writes to #norsktale-data.
  // Content.js (ISOLATED world) polls it. DOM is shared between worlds.
  let lastDataVersion = "";
  function pollGDocsData() {
    if (!isGoogleDocs()) return;
    if (replaceInProgress) {
      const el = document.getElementById("norsktale-data");
      const done = el && el.getAttribute("data-replace-done") === "true";
      const timedOut = Date.now() - replaceStartTime > 4000;
      if (done || timedOut) {
        norsktaleLog("Replace " + (done ? "confirmed" : "timed out") + " — resuming text updates");
        replaceInProgress = false;
        if (el) el.removeAttribute("data-replace-done");
      } else {
        return;
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
  setInterval(pollGDocsData, 500);

  function handleResponse(msg) {
    if (msg.action !== "replace") return;
    norsktaleLog("REPLACE: find='" + (msg.expected||"") + "' text='" + (msg.text||"") + "'");

    if (isGoogleDocs()) {
      const findText = msg.expected || "";
      const replaceText = msg.text || "";
      if (!findText) {
        norsktaleLog("REPLACE FAILED: no expected text");
        return;
      }
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
      replaceInProgress = true;
      replaceStartTime = Date.now();
      lastSent = "";
      lastDataVersion = "";
      return;
    }

    const el = activeElement || lastTextElement;
    if (!el) return;
    const replacement = msg.text;
    const expected = msg.expected || "";

    // msg.start / msg.end are paragraph-relative; convert to document-absolute
    let start = lastParaStart + (msg.start || 0);
    let end   = lastParaStart + (msg.end   || 0);

    if (el.tagName === "TEXTAREA" || el.tagName === "INPUT") {
      const val = el.value;
      if (expected && val.substring(start, end).toLowerCase() !== expected.toLowerCase()) {
        const idx = val.toLowerCase().indexOf(expected.toLowerCase(), Math.max(0, start - 5));
        if (idx >= 0 && idx <= start + 10) { start = idx; end = idx + expected.length; }
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
      lastParaStart = 0;
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

  // Converts viewport-relative coordinates to physical screen pixels.
  // Must match the physical pixel coordinates the native app uses (ClientToScreen on Windows).
  function viewportToScreen(vpX, vpY) {
    const dpr = window.devicePixelRatio || 1;
    const chromeHeight = window.outerHeight - window.innerHeight;
    return {
      x: Math.round((window.screenX + vpX) * dpr),
      y: Math.round((window.screenY + chromeHeight + vpY + 25) * dpr)
    };
  }

  function getTextareaCursorXY(el) {
    const mirror = document.createElement("div");
    const style = window.getComputedStyle(el);
    for (const prop of [
      "fontFamily","fontSize","fontWeight","fontStyle","letterSpacing","wordSpacing",
      "lineHeight","textTransform","paddingTop","paddingRight","paddingBottom","paddingLeft",
      "borderTopWidth","borderRightWidth","borderBottomWidth","borderLeftWidth",
      "boxSizing","width","whiteSpace","wordWrap","overflowWrap","tabSize"
    ]) {
      mirror.style[prop] = style[prop];
    }
    mirror.style.position = "absolute";
    mirror.style.left = "-9999px";
    mirror.style.top = "0";
    mirror.style.whiteSpace = "pre-wrap";
    mirror.style.wordWrap = "break-word";
    mirror.style.visibility = "hidden";
    mirror.style.overflow = "hidden";
    const textBefore = el.value.substring(0, el.selectionStart);
    mirror.textContent = textBefore;
    const marker = document.createElement("span");
    marker.textContent = "|";
    mirror.appendChild(marker);
    document.body.appendChild(mirror);
    const rect = el.getBoundingClientRect();
    const markerRect = marker.getBoundingClientRect();
    const mirrorRect = mirror.getBoundingClientRect();
    const vpX = rect.left + (markerRect.left - mirrorRect.left) - el.scrollLeft;
    const vpY = rect.top + (markerRect.top - mirrorRect.top) - el.scrollTop;
    document.body.removeChild(mirror);
    return viewportToScreen(vpX, vpY);
  }

  function getTextAndCursor(el) {
    if (isGoogleDocs()) return null;
    if (!el) return null;

    if (el.tagName === "TEXTAREA" || (el.tagName === "INPUT" && el.type === "text")) {
      const pos = getTextareaCursorXY(el);
      const { paraText, cursorInPara, paraStart } =
        extractParagraphAtCursor(el.value, el.selectionStart);
      lastParaStart = paraStart;
      return { text: paraText, cursorStart: cursorInPara, cursorEnd: cursorInPara, caretX: pos.x, caretY: pos.y };
    }

    if (el.isContentEditable) {
      const sel = window.getSelection();
      // innerText normalises Gmail's per-div line breaks to \n
      const fullText = el.innerText.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
      let docCursor = 0;
      let caretX = 0, caretY = 0;
      if (sel && sel.rangeCount > 0) {
        const range = sel.getRangeAt(0);
        const preRange = document.createRange();
        preRange.selectNodeContents(el);
        preRange.setEnd(range.startContainer, range.startOffset);
        docCursor = preRange.toString().length;
        const caretRect = range.getBoundingClientRect();
        if (caretRect && caretRect.height > 0) {
          const screenPos = viewportToScreen(caretRect.left, caretRect.bottom);
          caretX = screenPos.x;
          caretY = screenPos.y;
        }
      }
      const { paraText, cursorInPara, paraStart } =
        extractParagraphAtCursor(fullText, docCursor);
      lastParaStart = paraStart;
      if (!paraText.trim()) return null;
      return { text: paraText, cursorStart: cursorInPara, cursorEnd: cursorInPara, caretX, caretY };
    }

    return null;
  }

  function sendUpdate() {
    if (isGoogleDocs()) return;
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

  document.addEventListener("visibilitychange", () => {
    if (document.hidden) lastGDocsText = null;
  }, true);

  setInterval(() => {
    if (!port) connectPort();
    if (isGoogleDocs()) {
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
