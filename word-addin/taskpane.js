/* NorskTale Word Add-in bridge.
 *
 * ⚠️ DO NOT send full document text. NEVER hash full document.
 * Only send the current sentence. See plan for design rationale.
 *
 * Strategy:
 * - Typing: detect current sentence from paragraph, send sentence + word + cursor
 * - Cursor move: send sentence + word at new position
 * - Paste/undo: scan for changed sentence hashes, send only changed (TODO: step 5)
 * - Sentence split: detect period typed, send both new sentences (TODO: step 6)
 */

const BRIDGE_URL = "https://localhost:3000";
var statusEl;

Office.onReady(function (info) {
    statusEl = document.getElementById("status");

    if (info.host === Office.HostType.Word) {
        setStatus("Koblet til Word", "ok");

        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            onSelectionChanged
        );

        // Poll for reply commands from Rust app
        setInterval(pollReplies, 100);
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

// --- Step 1: Sentence detection ---

var SENTENCE_DELIMITERS = /[.!?:]/;

/**
 * Find the current sentence from paragraph text and cursor offset within it.
 * Returns { sentence, wordAtCursor, sentenceStartInPara }
 */
function detectSentence(paraText, cursorOffsetInPara) {
    // Find sentence start: scan backwards from cursor for delimiter
    var sentStart = 0;
    for (var i = cursorOffsetInPara - 1; i >= 0; i--) {
        if (SENTENCE_DELIMITERS.test(paraText[i])) {
            sentStart = i + 1;
            // Skip whitespace after delimiter
            while (sentStart < paraText.length && paraText[sentStart] === " ") {
                sentStart++;
            }
            break;
        }
    }

    // Find sentence end: scan forwards from cursor for delimiter or end of para
    var sentEnd = paraText.length;
    for (var i = cursorOffsetInPara; i < paraText.length; i++) {
        if (SENTENCE_DELIMITERS.test(paraText[i])) {
            sentEnd = i + 1; // include the delimiter
            break;
        }
    }

    var sentence = paraText.substring(sentStart, sentEnd).trim();

    // Find word at cursor
    var wordStart = cursorOffsetInPara;
    while (wordStart > 0 && isWordChar(paraText[wordStart - 1])) {
        wordStart--;
    }
    var wordEnd = cursorOffsetInPara;
    while (wordEnd < paraText.length && isWordChar(paraText[wordEnd])) {
        wordEnd++;
    }
    var wordAtCursor = paraText.substring(wordStart, wordEnd);

    return {
        sentence: sentence,
        wordAtCursor: wordAtCursor,
        sentenceStartInPara: sentStart
    };
}

function isWordChar(ch) {
    if (!ch) return false;
    var code = ch.charCodeAt(0);
    // a-z, A-Z, 0-9, æøåÆØÅ, hyphen, apostrophe
    return (code >= 65 && code <= 90) || (code >= 97 && code <= 122) ||
           (code >= 48 && code <= 57) || ch === "-" || ch === "'" ||
           ch === "æ" || ch === "ø" || ch === "å" ||
           ch === "Æ" || ch === "Ø" || ch === "Å";
}

// --- Event handler ---

var lastSentKey = "";

function onSelectionChanged() {
    Word.run(function (ctx) {
        var sel = ctx.document.getSelection();
        var para = sel.paragraphs.getFirst();
        // Get range from paragraph start to cursor to measure cursor offset in paragraph
        var paraRange = para.getRange("Start");
        var beforeCursor = paraRange.expandTo(sel.getRange("Start"));
        sel.load("start");
        para.load("text");
        beforeCursor.load("text");
        return ctx.sync().then(function () {
            var paraText = para.text;
            var cursorInPara = beforeCursor.text.length;

            if (cursorInPara > paraText.length) cursorInPara = paraText.length;

            var result = detectSentence(paraText, cursorInPara);

            // Dedup: don't send if nothing changed
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
                    sentenceStart: sel.start - cursorInPara + result.sentenceStartInPara
                })
            }).catch(function () {
                setStatus("Kan ikke nå NorskTale-app", "err");
            });
        });
    }).catch(function () {});
}

// --- Reply polling (kept from before) ---

function pollReplies() {
    fetch(BRIDGE_URL + "/reply")
        .then(function (resp) { return resp.json(); })
        .then(function (data) {
            if (!data || !data.action) return;
            if (data.action === "replace" && data.expected && data.text) {
                doReplace(data.expected, data.text);
            } else if (data.action === "replaceWord" && data.text) {
                doReplaceCurrentWord(data.text);
            }
        })
        .catch(function () {});
}

function doReplace(expected, replacement) {
    Word.run(function (ctx) {
        var results = ctx.document.body.search(expected, { matchCase: true });
        results.load("items");
        return ctx.sync().then(function () {
            if (results.items.length > 0) {
                results.items[0].insertText(replacement, "Replace");
                return ctx.sync();
            }
        });
    }).catch(function () {});
}

function doReplaceCurrentWord(replacement) {
    Word.run(function (ctx) {
        var sel = ctx.document.getSelection();
        var wordRange = sel.getRange("Whole");
        wordRange.insertText(replacement, "Replace");
        return ctx.sync();
    }).catch(function () {});
}
