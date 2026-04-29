/* NorskTale event-based activation handler.
 *
 * Runs automatically when a Word document opens (OnDocumentOpened).
 * Scans all paragraphs and POSTs them to the Rust desktop app so
 * spelling/grammar errors are detected immediately — even before
 * the user opens the taskpane.
 */

var BRIDGE_URL = "https://localhost:3000";

Office.onReady(function () {
    // Office.js is ready.
});

function hashString(str) {
    var hash = 0x811c9dc5;
    for (var i = 0; i < str.length; i++) {
        hash ^= str.charCodeAt(i);
        hash = (hash * 0x01000193) >>> 0;
    }
    return hash;
}

async function onDocumentOpened(event) {
    // Generate a document ID (same logic as taskpane.js)
    var documentId = (Office.context && Office.context.document && Office.context.document.url)
        || ("unsaved-" + Math.random().toString(36).substring(2, 10));

    try {
        // Tell the Rust app to reset state for this document
        await fetch(BRIDGE_URL + "/reset", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ documentName: documentId })
        });

        // Scan all paragraphs
        await Word.run(async function (ctx) {
            var paragraphs = ctx.document.body.paragraphs;
            paragraphs.load("items");
            await ctx.sync();

            var items = paragraphs.items;
            for (var i = 0; i < items.length; i++) {
                items[i].load("text,uniqueLocalId");
            }
            await ctx.sync();

            var changed = [];
            for (var i = 0; i < items.length; i++) {
                var paraText = items[i].text;
                if (paraText.trim().length < 2) continue;
                changed.push({
                    paragraphId: items[i].uniqueLocalId,
                    text: paraText
                });
            }

            if (changed.length > 0) {
                await fetch(BRIDGE_URL + "/changed", {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        type: "changed",
                        documentName: documentId,
                        paragraphs: changed
                    })
                });
            }

            // Tag document for auto-open taskpane on future opens
            try {
                Office.context.document.settings.set(
                    "Office.AutoShowTaskpaneWithDocument", true
                );
                Office.context.document.settings.saveAsync();
            } catch (e) { /* ignore if not supported */ }
        });
    } catch (e) {
        // App not running — silently fail. The taskpane retry will
        // handle it when the user eventually opens the taskpane.
    }

    // Required: signal that event processing is complete
    event.completed();
}

// Register the event handler — must match FunctionName in manifest
Office.actions.associate("onDocumentOpened", onDocumentOpened);
