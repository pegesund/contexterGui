// Spell content script v8 — active cursor required
// Google Docs: text from canvas hook (gdocs-inject.js), replace via postMessage
// Other sites: textarea/contenteditable direct
// RULE: NEVER send text unless user has a blinking cursor in an editable element.

(function() {
  "use strict";
  console.log("Spell content.js v8 loaded");

  function spellLog(msg) {
    console.log("Spell: " + msg);
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
      port = chrome.runtime.connect({ name: "spell-content" });
      port.onMessage.addListener(handleResponse);
      port.onDisconnect.addListener(() => {
        // Drain runtime.lastError so Chrome doesn't log
        // "Unchecked runtime.lastError" when the SW (or the page going
        // into bfcache) closes the port. The error is expected on
        // normal page lifecycle events — we don't need to act on it,
        // just acknowledge it.
        void chrome.runtime.lastError;
        port = null;
      });
    } catch (e) { port = null; }
  }

  function postToBackground(msg) {
    if (!port) connectPort();
    if (!port) return false;
    try {
      port.postMessage(msg);
      return true;
    } catch (e) {
      port = null;
      return false;
    }
  }

  // When the page is restored from bfcache, the old `port` reference is
  // dead but our code doesn't know until the next postMessage fails.
  // Proactively reconnect on pageshow so the very first textUpdate after
  // a forward/back navigation goes through cleanly.
  window.addEventListener("pageshow", (evt) => {
    if (evt.persisted) {
      // Page was restored from bfcache — drop the dead port and lazily
      // reconnect on the next message attempt (connectPort is called
      // from any !port check site).
      port = null;
    }
  });

  // When the user switches browser tabs the current tab's content script
  // keeps running, BUT its `lastSent` dedup string and the desktop's
  // `last_doc_text` snapshot diverge — the desktop still holds the OLD
  // tab's text, and the next textUpdate from this tab gets filtered by
  // lastSent if its content is identical to the previously-sent string
  // from this same tab.  Result reported 2026-05-19: "If I switch
  // between tabs inside browser, spell desktop app does not show errors
  // until I refresh the page."
  //
  // Fix: listen for visibility changes only (not every focus event —
  // window focus fires for things like Cmd+Tab back into Chrome while
  // the tab was already foreground in its window, and we don't want to
  // tear down the port for that). When the tab transitions FROM hidden
  // TO visible, drop the dedup key + port and kick a fresh send.
  let wasHidden = document.hidden;
  document.addEventListener("visibilitychange", () => {
    if (wasHidden && !document.hidden) {
      lastSent = "";
      lastDataVersion = "";
      // Drop the port too — Chrome sometimes pauses message delivery
      // while a tab is in the background, and a paused port can fail
      // silently when the tab returns.  Null here triggers a fresh
      // connectPort on the next sendUpdate / pollGDocsData call.
      port = null;
      // Kick a fresh send immediately so the user doesn't have to type
      // to see their existing errors come back.
      try { sendUpdate(); } catch(e) {}
      try { pollGDocsData(); } catch(e) {}
    }
    wasHidden = document.hidden;
  });

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

  // gdocs-inject.js (MAIN world) writes to #spell-data.
  // Content.js (ISOLATED world) polls it. DOM is shared between worlds.
  let lastDataVersion = "";

  function resumeGDocsUpdates(reason) {
    spellLog("Replace " + reason + " - resuming text updates");
    replaceInProgress = false;
    replaceStartTime = 0;
    lastSent = "";
    lastDataVersion = "";
    setTimeout(pollGDocsData, 0);
  }

  function handleGDocsReplaceDone(event) {
    const replEl = event.currentTarget;
    const succeeded = replEl && replEl.getAttribute("data-result") === "true";
    resumeGDocsUpdates(succeeded ? "confirmed" : "failed");
  }

  function ensureGDocsReplaceElement() {
    let replEl = document.getElementById("spell-replace");
    if (!replEl) {
      replEl = document.createElement("div");
      replEl.id = "spell-replace";
      replEl.style.display = "none";
      document.documentElement.appendChild(replEl);
    }
    replEl.addEventListener("spell-replace-done", handleGDocsReplaceDone);
    return replEl;
  }

  function pollGDocsData() {
    if (!isGoogleDocs()) return;
    if (replaceInProgress) {
      const timedOut = Date.now() - replaceStartTime > 4000;
      if (timedOut) {
        resumeGDocsUpdates("timed out");
      } else {
        return;
      }
    }
    const el = document.getElementById("spell-data");
    if (!el) return;
    const text = el.getAttribute("data-text");
    const selectedText = el.getAttribute("data-selected-text") || "";
    const cursor = parseInt(el.getAttribute("data-cursor") || "0", 10);
    const paragraphStart = parseInt(el.getAttribute("data-paragraph-start") || "0", 10);
    const caretX = parseInt(el.getAttribute("data-caret-x") || "0", 10);
    const caretY = parseInt(el.getAttribute("data-caret-y") || "0", 10);
    if (text === null) return;
    const version = text + "|" + cursor + "|" + paragraphStart + "|" + caretX + "|" + caretY + "|" + selectedText;
    if (version === lastDataVersion) return;
    const data = {
      text: text, cursorStart: cursor, cursorEnd: cursor,
      paragraphStart: paragraphStart, selectedText: selectedText,
    };
    lastGDocsText = data;
    const key = text + "|" + cursor + "|" + paragraphStart + "|" + selectedText;
    if (key === lastSent) return;
    spellLog("GDOCS canvas text: " + text.length + " chars, caret=(" + caretX + "," + caretY + ")");
    if (postToBackground({
      type: "textUpdate", text: text,
      cursorStart: cursor, cursorEnd: cursor,
      paragraphStart: paragraphStart,
      selectedText: selectedText,
      caretX: caretX, caretY: caretY, url: window.location.href
    })) {
      lastSent = key;
      lastDataVersion = version;
    }
  }
  setInterval(pollGDocsData, 500);

  function handleResponse(msg) {
    if (msg.action !== "replace") return;
    spellLog("REPLACE: find='" + (msg.expected||"") + "' text='" + (msg.text||"") + "'");

    if (isGoogleDocs()) {
      const findText = msg.expected || "";
      const replaceText = msg.text || "";
      if (!findText) {
        spellLog("REPLACE FAILED: no expected text");
        return;
      }
      const replEl = ensureGDocsReplaceElement();
      replEl.setAttribute("data-find", findText);
      replEl.setAttribute("data-replace", replaceText);
      replEl.setAttribute("data-offset", String(msg.start || 0));
      replEl.setAttribute("data-paragraph-start", String(msg.paragraphStart || 0));
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
      const selectedText = el.selectionEnd > el.selectionStart
        ? el.value.substring(el.selectionStart, el.selectionEnd)
        : "";
      return {
        text: paraText, cursorStart: cursorInPara, cursorEnd: cursorInPara,
        selectedText, caretX: pos.x, caretY: pos.y,
      };
    }

    if (el.isContentEditable) {
      const sel = window.getSelection();
      // innerText normalises Gmail's per-div line breaks to \n
      const fullText = el.innerText.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
      let docCursor = 0;
      let caretX = 0, caretY = 0;
      let selectedText = "";
      if (sel && sel.rangeCount > 0) {
        const range = sel.getRangeAt(0);
        const preRange = document.createRange();
        preRange.selectNodeContents(el);
        preRange.setEnd(range.startContainer, range.startOffset);
        docCursor = preRange.toString().length;
        if (!range.collapsed && el.contains(range.commonAncestorContainer)) {
          selectedText = range.toString();
        }
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
      // Previously: `if (!paraText.trim()) return null;` — that swallowed
      // the empty state entirely.  Result: after Cmd+A + Backspace in a
      // Gmail compose / Reddit comment box, the editor became empty but
      // we never told the desktop, so its writing_errors stayed showing
      // the old misspellings until the user started typing again.
      // Reported 2026-05-19.  Now we let empty text through; the
      // sendUpdate caller forwards it, the browser bridge stores "",
      // and the desktop's prune_resolved_errors() empty-doc clear at
      // main.rs:3335 fires.
      return {
        text: paraText, cursorStart: cursorInPara, cursorEnd: cursorInPara,
        selectedText, caretX, caretY,
      };
    }

    return null;
  }

  function sendUpdate() {
    if (isGoogleDocs()) return;
    const el = getElementWithActiveCursor();
    if (!el) return;
    const data = getTextAndCursor(el);
    // Forward the update even when data.text is empty — the desktop relies
    // on a `text:""` event to detect Cmd+A + Backspace and clear stale
    // errors via prune_resolved_errors's empty-doc branch.  Previous guard
    // (`if (!data || !data.text) return;`) dropped those sends so the user
    // saw old errors lingering over an empty comment box.
    if (!data) return;
    const key = data.text + "|" + data.cursorStart + "|" + data.selectedText;
    if (key === lastSent) return;
    activeElement = el;
    lastTextElement = el;
    if (el.spellcheck !== false) el.spellcheck = false;
    if (postToBackground({
      type: "textUpdate", text: data.text,
      cursorStart: data.cursorStart, cursorEnd: data.cursorEnd,
      selectedText: data.selectedText,
      caretX: data.caretX, caretY: data.caretY, url: window.location.href
    })) {
      lastSent = key;
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
    if (isGoogleDocs()) {
      if (replaceInProgress) {
        postToBackground({ type: "keepalive" });
        return;
      }
      if (document.hasFocus() && lastGDocsText) {
        postToBackground({
          type: "textUpdate", text: lastGDocsText.text,
          cursorStart: lastGDocsText.cursorStart, cursorEnd: lastGDocsText.cursorEnd,
          paragraphStart: lastGDocsText.paragraphStart,
          selectedText: lastGDocsText.selectedText,
          caretX: 0, caretY: 0, url: window.location.href
        });
      } else {
        postToBackground({ type: "keepalive" });
      }
      return;
    }
    const el = getElementWithActiveCursor();
    if (!el) return;
    const data = getTextAndCursor(el);
    // Allow empty text through — see comment in sendUpdate() above.
    if (!data) return;
    postToBackground({
      type: "textUpdate", text: data.text,
      cursorStart: data.cursorStart, cursorEnd: data.cursorEnd,
      selectedText: data.selectedText,
      caretX: data.caretX, caretY: data.caretY, url: window.location.href
    });
  }, 2000);
})();
