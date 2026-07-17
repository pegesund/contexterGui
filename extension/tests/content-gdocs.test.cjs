const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

class FakeElement {
  constructor(elements) {
    this.attributes = new Map();
    this.listeners = new Map();
    this.style = {};
    this.elements = elements;
    this._id = "";
  }

  get id() { return this._id; }
  set id(value) {
    this._id = value;
    if (value) this.elements.set(value, this);
  }

  setAttribute(name, value) { this.attributes.set(name, String(value)); }
  getAttribute(name) { return this.attributes.has(name) ? this.attributes.get(name) : null; }
  removeAttribute(name) { this.attributes.delete(name); }

  addEventListener(type, listener) {
    if (!this.listeners.has(type)) this.listeners.set(type, new Set());
    this.listeners.get(type).add(listener);
  }

  dispatchEvent(event) {
    event.currentTarget = this;
    event.target = this;
    for (const listener of this.listeners.get(event.type) || []) listener(event);
  }
}

function loadContentScript() {
  const elements = new Map();
  const intervals = [];
  const messages = [];
  let responseHandler = null;

  const port = {
    onMessage: { addListener(listener) { responseHandler = listener; } },
    onDisconnect: { addListener() {} },
    postMessage(message) { messages.push(message); },
  };
  const location = {
    hostname: "docs.google.com",
    pathname: "/document/d/test/edit",
    href: "https://docs.google.com/document/d/test/edit",
  };
  const document = {
    hidden: false,
    documentElement: { appendChild(element) { elements.set(element.id, element); } },
    addEventListener() {},
    createElement() { return new FakeElement(elements); },
    getElementById(id) { return elements.get(id) || null; },
    hasFocus() { return true; },
  };
  const context = {
    chrome: {
      runtime: {
        connect() { return port; },
        lastError: null,
      },
    },
    console: { log() {} },
    document,
    location,
    setInterval(callback, delay) {
      intervals.push({ callback, delay });
      return intervals.length;
    },
    setTimeout(callback) { callback(); },
    window: {
      addEventListener() {},
      location,
    },
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.join(__dirname, "..", "content.js"), "utf8");
  vm.runInContext(source, context, { filename: "content.js" });

  return {
    createData(text, cursor = 0, paragraphStart = 0) {
      const element = new FakeElement(elements);
      element.id = "spell-data";
      element.setAttribute("data-text", text);
      element.setAttribute("data-cursor", cursor);
      element.setAttribute("data-paragraph-start", paragraphStart);
      element.setAttribute("data-caret-x", 10);
      element.setAttribute("data-caret-y", 20);
      return element;
    },
    messages,
    poll: intervals.find(({ delay }) => delay === 500).callback,
    keepalive: intervals.find(({ delay }) => delay === 2000).callback,
    replacementElement() { return elements.get("spell-replace"); },
    respond(message) { responseHandler(message); },
  };
}

test("Google Docs replacement completion immediately resumes text updates", () => {
  const harness = loadContentScript();
  harness.createData("Jeg liker piza.", 14, 32);
  harness.poll();
  const update = harness.messages.find(({ type }) => type === "textUpdate");
  assert.equal(update.paragraphStart, 32);
  harness.respond({ action: "replace", expected: "piza", text: "pipa", start: 10, paragraphStart: 32 });

  harness.messages.length = 0;
  harness.keepalive();
  assert.equal(harness.messages.at(-1).type, "keepalive");

  const replacement = harness.replacementElement();
  assert.equal(replacement.getAttribute("data-paragraph-start"), "32");
  replacement.setAttribute("data-result", "true");
  replacement.dispatchEvent({ type: "spell-replace-done" });

  harness.messages.length = 0;
  harness.keepalive();
  assert.equal(harness.messages.at(-1).type, "textUpdate");
});

test("Google Docs forwards an empty active paragraph", () => {
  const harness = loadContentScript();
  harness.createData("", 0);
  harness.poll();

  const update = harness.messages.find(({ type }) => type === "textUpdate");
  assert.ok(update);
  assert.equal(update.text, "");
});
