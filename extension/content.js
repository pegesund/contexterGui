// NorskTale content script — monitors textareas and contenteditable elements
// Sends text + cursor position to the background script via native messaging

(function() {
  "use strict";
  console.log("NorskTale content.js v2 (text-node search) loaded");

  let lastSent = "";
  let lastCursor = -1;
  let activeElement = null;
  let lastTextElement = null; // persists even when focus leaves the textarea
  let port = null;

  function connectPort() {
    try {
      port = chrome.runtime.connect({ name: "norsktale-content" });
      port.onMessage.addListener(handleResponse);
      port.onDisconnect.addListener(() => { port = null; });
    } catch (e) {
      port = null;
    }
  }

  function handleResponse(msg) {
    console.log("NorskTale received:", JSON.stringify(msg));
    if (msg.action === "replace") {
      const el = activeElement || lastTextElement;
      console.log("NorskTale replace:", msg.text, "el:", el?.tagName, "active:", !!activeElement, "last:", !!lastTextElement);
      if (!el) { console.log("NorskTale: no element!"); return; }
      let start = msg.start;
      let end = msg.end;
      const replacement = msg.text;
      const expected = msg.expected || ""; // expected text at [start, end] for verification
      console.log("NorskTale: start:", start, "end:", end, "expected:", JSON.stringify(expected));
      if (el.tagName === "TEXTAREA" || el.tagName === "INPUT") {
        // Verify offset: check text at [start, end] matches expected
        const val = el.value;
        if (expected && val.substring(start, end).toLowerCase() !== expected.toLowerCase()) {
          console.log("NorskTale: offset mismatch! got:", JSON.stringify(val.substring(start, end)), "expected:", JSON.stringify(expected));
          // Search nearby (±5 chars) for the expected text
          const searchLower = expected.toLowerCase();
          const valLower = val.toLowerCase();
          const nearby = valLower.indexOf(searchLower, Math.max(0, start - 5));
          if (nearby >= 0 && nearby <= start + 5) {
            console.log("NorskTale: corrected offset", start, "→", nearby);
            start = nearby;
            end = nearby + expected.length;
          }
        }
        // Use native input setter to bypass React/framework issues
        const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLTextAreaElement.prototype, 'value')?.set
          || Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;
        const newVal = val.substring(0, start) + replacement + val.substring(end);
        if (nativeSetter) {
          nativeSetter.call(el, newVal);
        } else {
          el.value = newVal;
        }
        el.selectionStart = start + replacement.length;
        el.selectionEnd = start + replacement.length;
        el.dispatchEvent(new Event("input", { bubbles: true }));
        el.dispatchEvent(new Event("change", { bubbles: true }));
        console.log("NorskTale: value after:", JSON.stringify(el.value));
        lastSent = "";
      } else if (el.isContentEditable) {
        // Focus and select the range, then insert
        el.focus();
        const sel = window.getSelection();
        const range = document.createRange();

        // Search for expected text directly in text node content.
        // This avoids innerText vs DOM offset mismatches (block element \n counting differs).
        let startNode = null, startOfs = 0, endNode = null, endOfs = 0;
        let found = false;

        if (expected) {
          // Collect all text nodes and their cumulative offsets in raw DOM text
          const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
          let buf = "";
          const nodes = [];
          while (walker.nextNode()) {
            nodes.push({ node: walker.currentNode, off: buf.length });
            buf += walker.currentNode.textContent;
          }
          // Search for expected text near the hint offset
          const bufLower = buf.toLowerCase();
          const searchLower = expected.toLowerCase();
          let idx = bufLower.indexOf(searchLower, Math.max(0, start - 50));
          if (idx < 0) idx = bufLower.indexOf(searchLower);
          if (idx >= 0) {
            const endIdx = idx + expected.length;
            for (const n of nodes) {
              const nodeEnd = n.off + n.node.textContent.length;
              if (!startNode && idx < nodeEnd) {
                startNode = n.node; startOfs = idx - n.off;
              }
              if (endIdx <= nodeEnd) {
                endNode = n.node; endOfs = endIdx - n.off;
                break;
              }
            }
            if (startNode && endNode) found = true;
          }
          console.log("NorskTale CE: text-node search for", JSON.stringify(expected),
            "found:", found, "at buf idx:", idx, "DOM nodes:", nodes.length);
        }

        if (!found) {
          // Fallback: walkNodes with offset (old approach, for messages without expected)
          let charCount = 0;
          const blockTags = new Set(["DIV","P","BR","LI","TR","BLOCKQUOTE","H1","H2","H3","H4","H5","H6"]);
          function walkNodes(parent) {
            for (let child = parent.firstChild; child; child = child.nextSibling) {
              if (startNode && endNode) return;
              if (child.nodeType === Node.TEXT_NODE) {
                const nodeLen = child.textContent.length;
                if (!startNode && charCount + nodeLen >= start) {
                  startNode = child; startOfs = start - charCount;
                }
                if (!endNode && charCount + nodeLen >= end) {
                  endNode = child; endOfs = end - charCount;
                }
                charCount += nodeLen;
              } else if (child.nodeType === Node.ELEMENT_NODE) {
                if (child.tagName === "BR") { charCount += 1; }
                else {
                  const isBlock = blockTags.has(child.tagName);
                  if (isBlock && child !== parent.firstElementChild) { charCount += 1; }
                  walkNodes(child);
                }
              }
            }
          }
          walkNodes(el);
          if (startNode && endNode) found = true;
          console.log("NorskTale CE: walkNodes fallback, charCount:", charCount, "found:", found);
        }

        if (found) {
          range.setStart(startNode, startOfs);
          range.setEnd(endNode, endOfs);
          const selectedText = range.toString();
          console.log("NorskTale CE: selected:", JSON.stringify(selectedText), "replacing with:", JSON.stringify(replacement));
          sel.removeAllRanges();
          sel.addRange(range);
          const ok = document.execCommand("insertText", false, replacement);
          console.log("NorskTale CE: execCommand insertText result:", ok);
          if (!ok) {
            // Fallback: direct DOM manipulation
            range.deleteContents();
            range.insertNode(document.createTextNode(replacement));
            sel.collapseToEnd();
            el.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: replacement }));
          }
        } else {
          console.log("NorskTale CE: could not find text for replacement");
        }
        lastSent = "";
      }
    }
  }

  // Hidden mirror div for measuring caret position in textareas
  let mirror = null;
  // Browser chrome height (tab bar, address bar, bookmarks)
  function chromeOffsetY() {
    return window.outerHeight - window.innerHeight;
  }

  function getCaretCoords(el, pos) {
    // screenX/screenY are in physical pixels on Windows with DPI scaling
    // getBoundingClientRect returns CSS pixels — scale them to match
    const dpr = window.devicePixelRatio || 1;
    const ofsY = window.screenY + chromeOffsetY();
    const ofsX = window.screenX;

    if (el.isContentEditable) {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0) {
        const range = sel.getRangeAt(0).cloneRange();
        range.collapse(true);
        const rect = range.getBoundingClientRect();
        if (rect.width || rect.height) {
          return { x: Math.round(rect.left * dpr + ofsX), y: Math.round(rect.bottom * dpr + ofsY) };
        }
      }
      const elRect = el.getBoundingClientRect();
      return { x: Math.round(elRect.left * dpr + ofsX), y: Math.round((elRect.top + 20) * dpr + ofsY) };
    }

    // Textarea/input: use mirror div technique
    if (!mirror) {
      mirror = document.createElement("div");
      mirror.style.cssText = "position:absolute;top:0;left:-9999px;visibility:hidden;white-space:pre-wrap;word-wrap:break-word;";
      document.body.appendChild(mirror);
    }

    const style = getComputedStyle(el);
    const props = ["fontFamily","fontSize","fontWeight","fontStyle","letterSpacing","lineHeight",
                   "textTransform","wordSpacing","textIndent","paddingTop","paddingLeft","paddingRight","paddingBottom",
                   "borderTopWidth","borderLeftWidth","boxSizing","width"];
    props.forEach(p => mirror.style[p] = style[p]);
    mirror.style.overflowWrap = "break-word";
    mirror.innerHTML = "";

    const textBefore = el.value.substring(0, pos);
    const textNode = document.createTextNode(textBefore);
    const span = document.createElement("span");
    span.textContent = "|";
    mirror.appendChild(textNode);
    mirror.appendChild(span);

    const elRect = el.getBoundingClientRect();
    const spanRect = span.getBoundingClientRect();
    const mirrorRect = mirror.getBoundingClientRect();

    const lineH = parseFloat(style.fontSize) * 1.2;
    const x = elRect.left + (spanRect.left - mirrorRect.left) - el.scrollLeft;
    const y = elRect.top + (spanRect.top - mirrorRect.top) - el.scrollTop + lineH;

    return {
      x: Math.round(x * dpr + ofsX),
      y: Math.round(y * dpr + ofsY)
    };
  }

  function getTextAndCursor(el) {
    if (!el) return null;

    if (el.tagName === "TEXTAREA" || (el.tagName === "INPUT" && el.type === "text")) {
      let caretX = 0, caretY = 0;
      try { const c = getCaretCoords(el, el.selectionStart); caretX = c.x; caretY = c.y; } catch(e) {}
      return {
        text: el.value,
        cursorStart: el.selectionStart,
        cursorEnd: el.selectionEnd,
        caretX, caretY
      };
    }

    if (el.isContentEditable) {
      const sel = window.getSelection();
      const text = el.innerText;
      let cursorStart = 0;
      let cursorEnd = 0;
      if (sel && sel.rangeCount > 0) {
        const range = sel.getRangeAt(0);
        const preRange = document.createRange();
        preRange.selectNodeContents(el);
        preRange.setEnd(range.startContainer, range.startOffset);
        cursorStart = preRange.toString().length;
        preRange.setEnd(range.endContainer, range.endOffset);
        cursorEnd = preRange.toString().length;
      }
      let caretX = 0, caretY = 0;
      try { const c = getCaretCoords(el, cursorStart); caretX = c.x; caretY = c.y; } catch(e) {}
      return { text, cursorStart, cursorEnd, caretX, caretY };
    }

    return null;
  }

  function sendUpdate() {
    const el = document.activeElement;
    if (!el) return;

    const data = getTextAndCursor(el);
    if (!data || !data.text) return;

    // Only send if changed
    const key = data.text + "|" + data.cursorStart;
    if (key === lastSent) return;
    lastSent = key;
    activeElement = el;
    lastTextElement = el;
    // Disable browser's built-in spellcheck — NorskTale handles it
    if (el.spellcheck !== false) el.spellcheck = false;

    if (!port) connectPort();
    if (port) {
      try {
        port.postMessage({
          type: "textUpdate",
          text: data.text,
          cursorStart: data.cursorStart,
          cursorEnd: data.cursorEnd,
          caretX: data.caretX,
          caretY: data.caretY,
          url: window.location.href
        });
      } catch (e) {
        port = null;
      }
    }
  }

  // Monitor input, selection changes, and focus
  document.addEventListener("input", sendUpdate, true);
  document.addEventListener("selectionchange", sendUpdate, true);
  document.addEventListener("focusin", () => setTimeout(sendUpdate, 50), true);
  document.addEventListener("click", () => setTimeout(sendUpdate, 50), true);
  document.addEventListener("keyup", sendUpdate, true);

  // Periodic heartbeat — resend even when unchanged to keep the data file fresh.
  // BrowserBridge checks file mtime (10s window), so we must keep touching it.
  function sendHeartbeat() {
    const el = activeElement || lastTextElement || document.activeElement;
    if (!el) return;
    const data = getTextAndCursor(el);
    if (!data || !data.text) return;
    if (!port) connectPort();
    if (port) {
      try {
        port.postMessage({
          type: "textUpdate",
          text: data.text,
          cursorStart: data.cursorStart,
          cursorEnd: data.cursorEnd,
          caretX: data.caretX,
          caretY: data.caretY,
          url: window.location.href
        });
      } catch (e) {
        port = null;
      }
    }
  }
  setInterval(sendHeartbeat, 2000);
})();
