// NorskTale Google Docs Add-on
// Reads document text and cursor position, relays to sidebar for native messaging bridge

function onOpen() {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Start NorskTale', 'showSidebar')
    .addToUi();
}

function onInstall() {
  onOpen();
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('NorskTale')
    .setWidth(250);
  DocumentApp.getUi().showSidebar(html);
}

function getDocumentData() {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var text = body.getText();

  // Get cursor position (wrapped in try/catch — getCursor can throw)
  var cursorPos = text.length;
  try {
    var cursor = doc.getCursor();
    if (cursor) {
      var element = cursor.getElement();
      var offset = cursor.getOffset();
      cursorPos = getAbsoluteOffset(body, element, offset);
    }
  } catch (e) {
    // Cursor position unavailable, default to end of text
  }

  // Get selection if any
  var selStart = cursorPos;
  var selEnd = cursorPos;
  try {
    var selection = doc.getSelection();
    if (selection) {
      var ranges = selection.getRangeElements();
      if (ranges.length > 0) {
        var first = ranges[0];
        var last = ranges[ranges.length - 1];
        selStart = getAbsoluteOffset(body, first.getElement(),
          first.isPartial() ? first.getStartOffset() : 0);
        selEnd = getAbsoluteOffset(body, last.getElement(),
          last.isPartial() ? last.getEndOffsetInclusive() + 1 : last.getElement().asText().getText().length);
      }
    }
  } catch (e) {
    // Selection unavailable
  }

  return {
    text: text,
    cursorStart: selStart,
    cursorEnd: selEnd
  };
}

function getAbsoluteOffset(body, element, offset) {
  // Walk through body paragraphs to find absolute character offset
  var absOffset = 0;
  var numChildren = body.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    var childText = child.asText ? child.asText().getText() : child.getText ? child.getText() : "";

    // Check if element is within this child
    if (containsElement(child, element)) {
      // Find offset within this paragraph
      absOffset += getOffsetWithin(child, element, offset);
      return absOffset;
    }

    absOffset += childText.length;
    if (i < numChildren - 1) absOffset += 1; // newline between paragraphs
  }

  return absOffset;
}

function containsElement(parent, target) {
  if (parent === target) return true;
  // For Text elements inside paragraphs
  if (parent.getType() === DocumentApp.ElementType.PARAGRAPH ||
      parent.getType() === DocumentApp.ElementType.LIST_ITEM) {
    var numChildren = parent.getNumChildren();
    for (var i = 0; i < numChildren; i++) {
      if (parent.getChild(i) === target) return true;
    }
  }
  return false;
}

function getOffsetWithin(parent, target, offset) {
  if (parent === target) return offset;
  // Text element inside a paragraph
  var numChildren = parent.getNumChildren();
  var pos = 0;
  for (var i = 0; i < numChildren; i++) {
    var child = parent.getChild(i);
    if (child === target) {
      return pos + offset;
    }
    pos += child.asText ? child.asText().getText().length : 0;
  }
  return offset;
}

function applyReplacement(start, end, replacement) {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var text = body.getText();

  // Simple approach: use body.replaceText() with the expected text
  // The "expected" text is sent as the find parameter from content.js
  // But replaceText uses regex, so we also try offset-based replacement

  // Try offset-based replacement first
  var startInfo = findPositionInBody(body, start);
  var endInfo = findPositionInBody(body, end);

  if (startInfo && endInfo && startInfo.element && endInfo.element) {
    if (startInfo.element === endInfo.element) {
      // Same text element — direct edit
      var textEl = startInfo.element.editAsText();
      textEl.deleteText(startInfo.offset, endInfo.offset - 1);
      textEl.insertText(startInfo.offset, replacement);
      return "ok";
    }
  }

  // Fallback: extract the text at [start, end] and use body.replaceText()
  var oldText = text.substring(start, end);
  if (oldText.length > 0) {
    // Escape regex special chars
    var escaped = oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    body.replaceText(escaped, replacement);
    return "ok_regex";
  }

  return "failed";
}

function findPositionInBody(body, absOffset) {
  var pos = 0;
  var numChildren = body.getNumChildren();
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    var childText = child.asText ? child.asText().getText() : "";
    var childLen = childText.length;
    if (pos + childLen >= absOffset) {
      // Target is in this paragraph
      var localOffset = absOffset - pos;
      // Find the Text element
      var numSubChildren = child.getNumChildren();
      var subPos = 0;
      for (var j = 0; j < numSubChildren; j++) {
        var sub = child.getChild(j);
        var subText = sub.asText ? sub.asText().getText() : "";
        if (subPos + subText.length >= localOffset) {
          return { element: sub, offset: localOffset - subPos };
        }
        subPos += subText.length;
      }
      return { element: child, offset: localOffset };
    }
    pos += childLen + 1; // +1 for newline
  }
  return null;
}
