const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

class FakeElement {
  constructor({ text = "", rect = null, children = [] } = {}) {
    this.textContent = text;
    this.rect = rect;
    this.children = children;
    this.attributes = new Map();
    this.style = {};
    this.listeners = new Map();
    this.id = "";
  }

  querySelectorAll(selector) {
    return selector === ".kix-wordhtmlgenerator-word-node" ? this.children : [];
  }

  getBoundingClientRect() { return this.rect || { left: 0, right: 0, top: 0, bottom: 0, height: 0, width: 0 }; }
  closest() { return null; }
  setAttribute(name, value) { this.attributes.set(name, String(value)); }
  getAttribute(name) { return this.attributes.has(name) ? this.attributes.get(name) : null; }
  addEventListener(type, listener) { this.listeners.set(type, listener); }
  dispatchEvent(event) { this.listeners.get(event.type)?.(event); }
}

function paragraph(text, top) {
  return new FakeElement({
    rect: { left: 0, right: 600, top, bottom: top + 20, height: 20, width: 600 },
    children: [new FakeElement({ text, rect: { left: 0, right: 100, top, bottom: top + 20, height: 20, width: 100 } })],
  });
}

function loadInjector(paragraphs, caret, annotatedText = null) {
  const elements = new Map();
  const intervals = [];
  const document = {
    querySelector(selector) {
      if (selector === ".kix-cursor-caret") return caret;
      if (selector === "iframe.docs-texteventtarget-iframe") return {};
      return null;
    },
    querySelectorAll(selector) {
      return selector === ".kix-paragraphrenderer" ? paragraphs : [];
    },
    getElementById(id) { return elements.get(id) || null; },
    createElement() { return new FakeElement(); },
    documentElement: {
      appendChild(element) { elements.set(element.id, element); },
    },
    dispatchEvent() {},
    hasFocus() { return true; },
  };
  const context = {
    console: { log() {} },
    document,
    window: { devicePixelRatio: 1, screenX: 0, outerHeight: 900, innerHeight: 800 },
    Event: class Event { constructor(type) { this.type = type; } },
    InputEvent: class InputEvent { constructor(type) { this.type = type; } },
    setInterval(callback, delay) { intervals.push({ callback, delay }); return intervals.length; },
  };
  vm.createContext(context);
  if (annotatedText) {
    context.__annotatedText = annotatedText;
    vm.runInContext(
      "globalThis._docs_annotate_getAnnotatedText = async () => globalThis.__annotatedText;",
      context,
    );
  }
  const source = fs.readFileSync(path.join(__dirname, "..", "gdocs-inject.js"), "utf8");
  vm.runInContext(source, context, { filename: "gdocs-inject.js" });

  return {
    async emit() {
      const poll = intervals.find(({ delay }) => delay === 500);
      await poll.callback();
    },
    data() { return elements.get("spell-data"); },
  };
}

test("Google Docs keeps duplicate paragraphs distinct by DOM position", async () => {
  const duplicate = "Jeg liker piza.";
  const first = paragraph(duplicate, 0);
  const second = paragraph(duplicate, 30);
  const caret = new FakeElement({ rect: { left: 100, right: 102, top: 30, bottom: 50, height: 20, width: 2 } });
  const injector = loadInjector([first, second], caret);

  await injector.emit();

  const data = injector.data();
  assert.equal(data.getAttribute("data-text"), duplicate);
  assert.equal(data.getAttribute("data-cursor"), String(duplicate.length));
  assert.equal(data.getAttribute("data-paragraph-start"), String(duplicate.length + 1));
});

test("Google Docs publishes the exact annotated selection for TTS", async () => {
  const text = "Jeg liker piza.";
  const para = paragraph(text, 0);
  const caret = new FakeElement({ rect: { left: 100, right: 102, top: 0, bottom: 20, height: 20, width: 2 } });
  const annotatedText = {
    getText() { return text; },
    getSelection() { return [{ start: 4, end: 14 }]; },
  };
  const injector = loadInjector([para], caret, annotatedText);

  await injector.emit();

  assert.equal(injector.data().getAttribute("data-selected-text"), "liker piza");
});
