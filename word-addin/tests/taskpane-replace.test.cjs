const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");

test("Word spelling replacement refreshes the paragraph before posting its echo", () => {
  const source = fs.readFileSync(path.join(__dirname, "..", "taskpane.js"), "utf8");
  const replaceStart = source.indexOf("function doReplace(expected, replacement, paragraphId)");
  const replaceEnd = source.indexOf("function doSelectWord", replaceStart);
  const replaceBody = source.slice(replaceStart, replaceEnd);

  assert.ok(replaceStart >= 0 && replaceEnd > replaceStart);
  const selectionEnd = replaceBody.indexOf('newRange.select("End");');
  const refreshedText = replaceBody.indexOf('scope.load("text");', selectionEnd);
  const sentReplacement = replaceBody.indexOf('spellReplacement: true');
  assert.ok(selectionEnd >= 0);
  assert.ok(refreshedText > selectionEnd);
  assert.ok(sentReplacement > refreshedText);
  assert.match(replaceBody, /text: updatedText,[\s\S]*spellReplacement: true/);
});
