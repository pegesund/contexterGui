/* NorskTale Word Add-in bridge.
 *
 * ⚠️ DO NOT send full document text. NEVER hash full document.
 * Send changed PARAGRAPHS only. Rust side handles sentence splitting + hashing.
 *
 * Strategy:
 * - Typing: detect current sentence from paragraph, send sentence + word + cursor
 * - Paragraph changed/added: send full paragraph text (Rust splits into sentences)
 * - Paragraph deleted: send paragraph IDs (Rust clears errors)
 * - Track: paragraphMap[uniqueLocalId] = hash of paragraph text
 */

var BRIDGE_URL = "https://localhost:3000";
var statusEl;
var SENTENCE_DELIMITERS = /[.!?:]/;
var lastSentKey = "";
var lastCursorParaId = "";
var lastCursorInPara = 0;
// Unique document identifier — URL for saved docs, random ID for unsaved
var documentId = (Office.context && Office.context.document && Office.context.document.url)
    || ("unsaved-" + Math.random().toString(36).substring(2, 10));

// Paragraph tracking: paragraphId -> hash of full paragraph text
var paragraphMap = {};

Office.onReady(function (info) {
    statusEl = document.getElementById("status");

    if (info.host === Office.HostType.Word) {
        setStatus("Koblet til Word", "ok");

        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            onSelectionChanged
        );

        setStatus("Starter skanning...", "ok");
        initialScan();

        setInterval(pollReplies, 100);

        // Light check every 5s: if paragraph count changed, rescan
        setInterval(checkParagraphCount, 5000);
    } else {
        setStatus("Ikke Word", "err");
    }
});

function setStatus(msg, cls) {
    if (statusEl) {
        statusEl.textContent = msg;
        statusEl.className = cls || "";
    }
}

// ── Hashing ──

function hashString(str) {
    var hash = 0x811c9dc5;
    for (var i = 0; i < str.length; i++) {
        hash ^= str.charCodeAt(i);
        hash = (hash * 0x01000193) >>> 0;
    }
    return hash;
}

// ── Initial scan: hash all paragraphs, send all to Rust ──

function initialScan() {
    // Clear all old errors on Rust side (new document or reload)
    var docName = documentId;
    fetch(BRIDGE_URL + "/reset", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ documentName: docName })
    }).catch(function () {});
    paragraphMap = {};

    enqueueWordRun(function () { return Word.run(function (ctx) {
        // Disable Word's built-in proofing — NorskTale handles spelling/grammar
        try {
            var bodyRange = ctx.document.body.getRange();
            bodyRange.hasNoProofing = true;
        } catch(e) { /* ignore if not supported */ }

        var paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        return ctx.sync().then(function () {
            var items = paragraphs.items;
            for (var i = 0; i < items.length; i++) {
                items[i].load("text,uniqueLocalId");
            }
            return ctx.sync().then(function () {
                var changed = [];
                for (var i = 0; i < items.length; i++) {
                    var paraId = items[i].uniqueLocalId;
                    var paraText = items[i].text;
                    if (paraText.trim().length < 2) continue; // skip empty
                    var h = hashString(paraText);
                    paragraphMap[paraId] = h;
                    changed.push({ paragraphId: paraId, text: paraText });
                }

                setStatus("Skannet " + items.length + " avsnitt, sender " + changed.length, "ok");

                if (changed.length > 0) {
                    sendChangedParagraphs(changed);
                }

                registerParagraphEvents(ctx);
                return ctx.sync();
            });
        });
    }).catch(function (err) {
        setStatus("Skann feilet: " + (err.message || String(err)), "err");
    }); });
}

// ── Paragraph events ──

function registerParagraphEvents(ctx) {
    ctx.document.onParagraphChanged.add(onParagraphChanged);
    ctx.document.onParagraphAdded.add(onParagraphAdded);
    ctx.document.onParagraphDeleted.add(onParagraphDeleted);
}

function onParagraphChanged(event) {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var ids = event.uniqueLocalIds;
        if (!ids || ids.length === 0) return ctx.sync();

        var paragraphs = [];
        for (var i = 0; i < ids.length; i++) {
            var para = ctx.document.getParagraphByUniqueLocalId(ids[i]);
            para.load("text,uniqueLocalId");
            paragraphs.push(para);
        }
        return ctx.sync().then(function () {
            var changed = [];
            for (var i = 0; i < paragraphs.length; i++) {
                var paraId = paragraphs[i].uniqueLocalId;
                var paraText = paragraphs[i].text;
                var newHash = hashString(paraText);
                var oldHash = paragraphMap[paraId];

                if (oldHash !== newHash) {
                    paragraphMap[paraId] = newHash;
                    changed.push({ paragraphId: paraId, text: paraText });
                }
            }
            if (changed.length > 0) {
                sendChangedParagraphs(changed);
            }
        });
    }).catch(function () {}); });
}

function onParagraphAdded(event) {
    setStatus("Para added — rescanning...", "ok");
    rescanAll();
}

function onParagraphDeleted(event) {
    setStatus("Para deleted — rescanning...", "ok");
    rescanAll();
}

/// Smart rescan: compare all current paragraphs against paragraphMap.
/// Only sends changed/new paragraphs. Detects deleted paragraphs.
function rescanAll() {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        return ctx.sync().then(function () {
            var items = paragraphs.items;
            for (var i = 0; i < items.length; i++) {
                items[i].load("text,uniqueLocalId");
            }
            return ctx.sync().then(function () {
                var changed = [];
                var currentIds = {};
                var deletedIds = [];

                for (var i = 0; i < items.length; i++) {
                    var paraId = items[i].uniqueLocalId;
                    var paraText = items[i].text;
                    currentIds[paraId] = true;

                    if (paraText.trim().length < 2) {
                        // Empty now — if it WAS in map, treat as deleted
                        if (paragraphMap[paraId] !== undefined) {
                            deletedIds.push(paraId);
                            delete paragraphMap[paraId];
                        }
                        continue;
                    }

                    var newHash = hashString(paraText);
                    var oldHash = paragraphMap[paraId];

                    if (oldHash !== newHash) {
                        // New or changed paragraph
                        paragraphMap[paraId] = newHash;
                        changed.push({ paragraphId: paraId, text: paraText });
                    }
                }

                // Find deleted paragraphs (in paragraphMap but not in current doc)
                for (var id in paragraphMap) {
                    if (!currentIds[id]) {
                        deletedIds.push(id);
                        delete paragraphMap[id];
                    }
                }

                var mapSize = Object.keys(paragraphMap).length;
                setStatus("Rescan: " + changed.length + " endret, " + deletedIds.length + " slettet (doc=" + items.length + " map=" + mapSize + ")", "ok");

                if (changed.length > 0) {
                    sendChangedParagraphs(changed);
                }
                if (deletedIds.length > 0) {
                    fetch(BRIDGE_URL + "/deleted", {
                        method: "POST",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ paragraphIds: deletedIds })
                    }).catch(function () {});
                }
            });
        });
    }).catch(function () {}); });
}

// Light paragraph count check — only loads count, not text
function checkParagraphCount() {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var paragraphs = ctx.document.body.paragraphs;
        paragraphs.load("items");
        return ctx.sync().then(function () {
            var currentCount = paragraphs.items.length;
            var knownCount = Object.keys(paragraphMap).length;
            if (currentCount !== knownCount) {
                rescanAll();
            }
        });
    }).catch(function () {}); });
}

// Debounced rescan — waits 1 second after last trigger to avoid repeated rescans
var rescanTimer = null;
function scheduleRescan() {
    if (rescanTimer) clearTimeout(rescanTimer);
    rescanTimer = setTimeout(function () {
        rescanTimer = null;
        rescanAll();
    }, 1000);
}

var totalSent = 0;
function sendChangedParagraphs(changed) {
    totalSent += changed.length;
    fetch(BRIDGE_URL + "/changed", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
            type: "changed",
            documentName: documentId,
            paragraphs: changed
        })
    }).then(function() {
        setStatus("Sendt " + totalSent + " avsnitt", "ok");
    }).catch(function (err) {
        setStatus("Feil: " + (err.message || err), "err");
    });
}

// ── Typing / cursor move ──

var selectionTimer = null; // debounce timer — one Word.run per pause
var lastSelStart = -1; // track cursor position
var lastSentWord = ""; // track last word sent — reject stale empty-word reads

function onSelectionChanged() {
    // Debounce: wait 80ms after last keystroke, then do ONE Word.run
    if (selectionTimer) clearTimeout(selectionTimer);
    selectionTimer = setTimeout(doSelectionRead, 80);
}

function doSelectionRead() {
    selectionTimer = null;
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var sel = ctx.document.getSelection();
        var para = sel.paragraphs.getFirst();
        var paraRange = para.getRange("Start");
        var beforeCursor = paraRange.expandTo(sel.getRange("Start"));
        sel.load("start");
        para.load("text,uniqueLocalId");
        beforeCursor.load("text");
        return ctx.sync().then(function () {
            var paraText = para.text;
            var cursorInPara = beforeCursor.text.length;
            if (cursorInPara > paraText.length) cursorInPara = paraText.length;

            // Check if paragraph is new or changed (paste/cut/drag) — trigger rescan
            var paraId = para.uniqueLocalId;
            var currentHash = hashString(paraText);
            if (paragraphMap[paraId] === undefined || paragraphMap[paraId] !== currentHash) {
                if (paragraphMap[paraId] !== currentHash) {
                    paragraphMap[paraId] = currentHash;
                    sendChangedParagraphs([{ paragraphId: paraId, text: paraText, cursorStart: cursorInPara }]);
                }
                if (paragraphMap[paraId] === undefined) {
                    scheduleRescan();
                }
            }

            // Store for fast replaceAtCursor
            lastCursorParaId = paraId;
            lastCursorInPara = cursorInPara;

            var result = detectSentence(paraText, cursorInPara);

            // Reject stale reads: empty word near a recent non-empty word position
            if (result.wordAtCursor === "" && lastSentWord !== ""
                && Math.abs(sel.start - lastSelStart) < 5) {
                return;
            }
            lastSelStart = sel.start;
            lastSentWord = result.wordAtCursor;

            var key = result.sentence + "|" + result.wordAtCursor + "|" + sel.start;
            if (key === lastSentKey) return;
            lastSentKey = key;

            setStatus("pos " + sel.start + " ord: " + result.wordAtCursor, "ok");

            return fetch(BRIDGE_URL + "/context", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    type: "typing",
                    sentence: result.sentence,
                    word: result.wordAtCursor,
                    cursorStart: sel.start,
                    paragraphId: para.uniqueLocalId,
                    documentName: documentId
                })
            }).then(function (resp) {
                return resp.json();
            }).then(function (data) {
                if (data && data.status === "rescan") {
                    setStatus("Byttet dokument — skanner...", "ok");
                    initialScan();
                }
            }).catch(function () {
                setStatus("Kan ikke nå NorskTale-app", "err");
            });
        });
    }).catch(function () {}); });
}

function detectSentence(paraText, cursorOffsetInPara) {
    var sentStart = 0;
    for (var i = cursorOffsetInPara - 1; i >= 0; i--) {
        if (SENTENCE_DELIMITERS.test(paraText[i])) {
            sentStart = i + 1;
            while (sentStart < paraText.length && paraText[sentStart] === " ") sentStart++;
            break;
        }
    }

    var sentEnd = paraText.length;
    for (var i = cursorOffsetInPara; i < paraText.length; i++) {
        if (SENTENCE_DELIMITERS.test(paraText[i])) {
            sentEnd = i + 1;
            break;
        }
    }

    var sentence = paraText.substring(sentStart, sentEnd).trim();

    var wordStart = cursorOffsetInPara;
    while (wordStart > 0 && isWordChar(paraText[wordStart - 1])) wordStart--;
    var wordEnd = cursorOffsetInPara;
    while (wordEnd < paraText.length && isWordChar(paraText[wordEnd])) wordEnd++;
    var wordAtCursor = paraText.substring(wordStart, wordEnd);

    return { sentence: sentence, wordAtCursor: wordAtCursor };
}

function isWordChar(ch) {
    if (!ch) return false;
    var code = ch.charCodeAt(0);
    return (code >= 65 && code <= 90) || (code >= 97 && code <= 122) ||
           (code >= 48 && code <= 57) || ch === "-" || ch === "'" ||
           ch === "æ" || ch === "ø" || ch === "å" ||
           ch === "Æ" || ch === "Ø" || ch === "Å";
}

// ── Reply polling ──

function pollReplies() {
    fetch(BRIDGE_URL + "/reply?doc=" + encodeURIComponent(documentId))
        .then(function (resp) { return resp.json(); })
        .then(function (data) {
            if (!data || !data.action) return;
            if (data.action === "replace" && data.expected && data.text) {
                doReplace(data.expected, data.text, data.paragraphId);
            } else if (data.action === "replaceWord" && data.text) {
                // text format: "prefix|replacement" — find prefix in paragraph, replace with replacement
                var parts = data.text.split("|");
                if (parts.length === 2) {
                    doReplaceAtCursor(parts[0], parts[1]);
                } else {
                    doReplaceAtCursor(data.text, data.text);
                }
            } else if (data.action === "underline" && data.word) {
                doUnderline(data.word, data.paragraphId, data.color || "red");
            } else if (data.action === "clearParagraphUnderlines" && data.paragraphId) {
                doClearParagraphUnderlines(data.paragraphId);
            } else if (data.action === "clearUnderline" && data.word) {
                doClearUnderline(data.word, data.paragraphId);
            } else if (data.action === "clearAllUnderlines") {
                doClearAllUnderlines();
            } else if (data.action === "replaceAtCursor" && data.expected && data.text) {
                doReplaceAtCursor(data.expected, data.text);
            } else if (data.action === "rescan") {
                paragraphMap = {};
                initialScan();
            } else if (data.action === "appendParagraph" && data.text) {
                doAppendParagraph(data.text);
            } else if (data.action === "deleteAfter" && data.text) {
                doDeleteAfter(data.text);
            } else if (data.action === "deleteText" && data.text) {
                doDeleteText(data.text);
            } else if (data.action === "selectWord" && data.word) {
                doSelectWord(data.word, data.paragraphId);
            }
        })
        .catch(function () {});
}

function doReplace(expected, replacement, paragraphId) {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        if (paragraphId) {
            var para = ctx.document.getParagraphByUniqueLocalId(paragraphId);
            var results = para.search(expected, { matchCase: false });
            results.load("items");
            return ctx.sync().then(function () {
                if (results.items.length > 0) {
                    results.items[0].insertText(replacement, "Replace");
                    return ctx.sync().then(function () {
                        // Trigger rescan so changed paragraph is detected
                        rescanAll();
                    });
                }
            });
        } else {
            var results = ctx.document.body.search(expected, { matchCase: false });
            results.load("items");
            return ctx.sync().then(function () {
                if (results.items.length > 0) {
                    results.items[0].insertText(replacement, "Replace");
                    return ctx.sync().then(function () {
                        // Trigger rescan so changed paragraph is detected
                        rescanAll();
                    });
                }
            });
        }
    }).catch(function () {}); });
}

function doSelectWord(word, paragraphId) {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var searchScope;
        if (paragraphId) {
            searchScope = ctx.document.getParagraphByUniqueLocalId(paragraphId);
        } else {
            searchScope = ctx.document.body;
        }
        var results = searchScope.search(word, { matchCase: false, matchWholeWord: true });
        results.load("items");
        return ctx.sync().then(function () {
            if (results.items.length > 0) {
                results.items[0].select();
                return ctx.sync();
            }
        });
    }).catch(function (e) { console.log("selectWord error:", e); }); });
}

var wordRunQueue = [];
var wordRunBusy = false;

function enqueueWordRun(fn) {
    wordRunQueue.push(fn);
    if (!wordRunBusy) drainWordRunQueue();
}

function drainWordRunQueue() {
    if (wordRunQueue.length === 0) { wordRunBusy = false; return; }
    wordRunBusy = true;
    var fn = wordRunQueue.shift();
    fn().then(function () { wordRunBusy = false; drainWordRunQueue(); }).catch(function () { wordRunBusy = false; drainWordRunQueue(); });
}

// Guard for event-driven Word.run calls — skip if queue is busy
function isWordBusy() { return wordRunBusy || wordRunQueue.length > 0; }

function doUnderline(word, paragraphId, color) {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var searchScope;
        if (paragraphId) {
            try {
                searchScope = ctx.document.getParagraphByUniqueLocalId(paragraphId);
            } catch(e) {
                searchScope = ctx.document.body;
            }
        } else {
            searchScope = ctx.document.body;
        }
        var results = searchScope.search(word, { matchCase: false, matchWholeWord: true });
        results.load("items/font");
        return ctx.sync().then(function () {
            if (results.items.length > 0) {
                results.items[0].font.underline = "Wave";
                try { results.items[0].font.underlineColor = color || "#FF0000"; } catch(e) {}
                fetch(BRIDGE_URL + "/log", { method: "POST", headers: {"Content-Type":"application/json"},
                    body: JSON.stringify({msg: "UNDERLINE OK: '" + word + "' matches=" + results.items.length})
                }).catch(function(){});
            } else {
                fetch(BRIDGE_URL + "/log", { method: "POST", headers: {"Content-Type":"application/json"},
                    body: JSON.stringify({msg: "UNDERLINE MISS: '" + word + "' not found in para=" + (paragraphId || "body")})
                }).catch(function(){});
            }
            return ctx.sync();
        });
    }).catch(function (e) {
        fetch(BRIDGE_URL + "/log", { method: "POST", headers: {"Content-Type":"application/json"},
            body: JSON.stringify({msg: "UNDERLINE ERROR: '" + word + "' " + e})
        }).catch(function(){});
    }); });
}

function doClearParagraphUnderlines(paragraphId) {
    if (!paragraphId) return;
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var para;
        try { para = ctx.document.getParagraphByUniqueLocalId(paragraphId); } catch(e) { return ctx.sync(); }
        var range = para.getRange("Whole");
        range.load("font");
        return ctx.sync().then(function () {
            if (range.font.underline === "Wave" || range.font.underline === "Mixed") {
                range.font.underline = "None";
            }
            return ctx.sync();
        });
    }).catch(function () {}); });
}

function doClearUnderline(word, paragraphId) {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var searchScope;
        if (paragraphId) {
            try { searchScope = ctx.document.getParagraphByUniqueLocalId(paragraphId); } catch(e) { searchScope = ctx.document.body; }
        } else {
            searchScope = ctx.document.body;
        }
        var results = searchScope.search(word, { matchCase: false, matchWholeWord: true });
        results.load("items/font");
        return ctx.sync().then(function () {
            for (var i = 0; i < results.items.length; i++) {
                results.items[i].font.underline = "None";
            }
            return ctx.sync();
        });
    }).catch(function () {}); });
}

function doAppendParagraph(text) {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        ctx.document.body.insertParagraph(text, "End");
        return ctx.sync();
    }).catch(function (e) { console.log("appendParagraph error:", e); }); });
}

function doDeleteAfter(marker) {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var body = ctx.document.body;
        var results = body.search(marker, { matchCase: false });
        results.load("items");
        return ctx.sync().then(function () {
            if (results.items.length > 0) {
                var last = results.items[results.items.length - 1];
                var rangeAfter = last.getRange("After");
                var endRange = body.getRange("End");
                var toDelete = rangeAfter.expandTo(endRange);
                toDelete.delete();
                return ctx.sync().then(function () {
                    // Add empty paragraph at end for clean test separation
                    body.insertParagraph("", "End");
                    return ctx.sync();
                });
            }
        });
    }).catch(function (e) { console.log("deleteAfter error:", e); }); });
}

function doDeleteText(text) {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var results = ctx.document.body.search(text, { matchCase: false });
        results.load("items");
        return ctx.sync().then(function () {
            for (var i = 0; i < results.items.length; i++) {
                results.items[i].delete();
            }
            return ctx.sync();
        });
    }).catch(function (e) { console.log("deleteText error:", e); }); });
}

function doClearAllUnderlines() {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var body = ctx.document.body;
        var range = body.getRange();
        range.font.underline = "None";
        return ctx.sync();
    }).catch(function (e) { console.log("clearAllUnderlines error:", e); }); });
}

function doReplaceAtCursor(prefix, replacement) {
    fetch(BRIDGE_URL + "/log", { method: "POST", headers: {"Content-Type":"application/json"},
        body: JSON.stringify({msg: "doReplaceAtCursor ENTER: prefix='" + prefix + "' replacement='" + replacement + "' paraId=" + lastCursorParaId + " cursor=" + lastCursorInPara})
    }).catch(function(){});
    if (!prefix) {
        // No prefix — insert at cursor using paragraph rewrite (same as non-empty prefix path)
        enqueueWordRun(function () { return Word.run(function (ctx) {
            var para = ctx.document.getSelection().paragraphs.getFirst();
            para.load("text");
            return ctx.sync().then(function () {
                var text = para.text;
                var pos = (cursorPos !== undefined && cursorPos <= text.length) ? cursorPos : text.length;
                var before = text.substring(0, pos);
                var after = text.substring(pos);
                var space = (after.length === 0 || (after[0] !== " " && after[0] !== "." && after[0] !== ",")) ? " " : "";
                var newText = before + replacement + space + after;
                fetch(BRIDGE_URL + "/log", { method: "POST", headers: {"Content-Type":"application/json"},
                    body: JSON.stringify({msg: "INSERT: from='" + text + "' to='" + newText + "' replacement='" + replacement + "' at pos=" + pos})
                }).catch(function(){});
                var cursorTarget = before.length + replacement.length + space.length;
                var inserted = para.insertText(newText, "Replace");
                para.load("uniqueLocalId");
                return ctx.sync().then(function () {
                    inserted.select("End");
                    var paraId = para.uniqueLocalId || "";
                    paragraphMap[paraId] = hashString(newText);
                    lastCursorInPara = cursorTarget;
                    lastCursorParaId = paraId;
                    sendChangedParagraphs([{ paragraphId: paraId, text: newText }]);
                    return ctx.sync();
                });
            });
        }).catch(function (e) {
            fetch(BRIDGE_URL + "/log", { method: "POST", headers: {"Content-Type":"application/json"},
                body: JSON.stringify({msg: "INSERT ERROR: " + e})
            }).catch(function(){});
        }); });
        return;
    }
    // Find prefix in paragraph text and replace it
    var cursorPos = lastCursorInPara || 0;

    enqueueWordRun(function () { return Word.run(function (ctx) {
        var para = ctx.document.getSelection().paragraphs.getFirst();
        para.load("text");
        return ctx.sync().then(function () {
            var text = para.text;
            // Find the occurrence of prefix closest to cursor position
            var bestPos = -1;
            var bestDist = 999999;
            var searchFrom = 0;
            while (true) {
                var pos = text.indexOf(prefix, searchFrom);
                if (pos < 0) break;
                // Only match at word boundaries (not inside other words)
                var before_ok = (pos === 0 || !isWordChar(text[pos - 1]));
                var after_ok = (pos + prefix.length >= text.length || !isWordChar(text[pos + prefix.length]));
                if (before_ok && after_ok) {
                    var dist = Math.abs(pos - cursorPos);
                    if (dist < bestDist) { bestDist = dist; bestPos = pos; }
                }
                searchFrom = pos + 1;
            }
            if (bestPos < 0) {
                fetch(BRIDGE_URL + "/log", { method: "POST", headers: {"Content-Type":"application/json"},
                    body: JSON.stringify({msg: "REPLACE FAIL: prefix='" + prefix + "' not found in text='" + text + "' cursor=" + cursorPos})
                }).catch(function(){});
                return ctx.sync();
            }
            var before = text.substring(0, bestPos);
            var after = text.substring(bestPos + prefix.length);
            var space = (after.length === 0 || (after[0] !== " " && after[0] !== "." && after[0] !== ",")) ? " " : "";
            var newText = before + replacement + space + after;
            fetch(BRIDGE_URL + "/log", { method: "POST", headers: {"Content-Type":"application/json"},
                body: JSON.stringify({msg: "REPLACE OK: prefix='" + prefix + "' → '" + replacement + "' at pos=" + bestPos + " before='" + text + "' after='" + newText + "'"})
            }).catch(function(){});
            var cursorTarget = before.length + replacement.length + space.length;
            var inserted = para.insertText(newText, "Replace");
            para.load("uniqueLocalId");
            return ctx.sync().then(function () {
                inserted.select("End");
                var paraId = para.uniqueLocalId || "";
                paragraphMap[paraId] = hashString(newText);
                lastCursorInPara = cursorTarget;
                lastCursorParaId = paraId;
                sendChangedParagraphs([{ paragraphId: paraId, text: newText }]);
                return ctx.sync();
            });
        });
    }).catch(function (e) {
        fetch(BRIDGE_URL + "/log", { method: "POST", headers: {"Content-Type":"application/json"},
            body: JSON.stringify({msg: "REPLACE ERROR: " + e})
        }).catch(function(){});
    }); });
}

function doReplaceCurrentWord(replacement) {
    enqueueWordRun(function () { return Word.run(function (ctx) {
        var sel = ctx.document.getSelection();
        var wordRange = sel.getRange("Whole");
        wordRange.insertText(replacement, "Replace");
        return ctx.sync();
    }).catch(function () {}); });
}
