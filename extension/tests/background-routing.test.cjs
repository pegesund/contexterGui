const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

class EventHook {
  constructor() { this.listeners = []; }
  addListener(listener) { this.listeners.push(listener); }
  emit(...args) { for (const listener of this.listeners) listener(...args); }
}

function createContentPort(tabId, frameId) {
  const incoming = new EventHook();
  const disconnect = new EventHook();
  return {
    name: "spell-content",
    sender: { tab: { id: tabId }, frameId },
    onMessage: incoming,
    onDisconnect: disconnect,
    sent: [],
    postMessage(message) { this.sent.push(message); },
    send(message) { incoming.emit(message); },
  };
}

function loadBackground() {
  const onConnect = new EventHook();
  const nativeMessages = new EventHook();
  const nativePort = {
    onMessage: nativeMessages,
    onDisconnect: new EventHook(),
    sent: [],
    postMessage(message) { this.sent.push(message); },
    disconnect() {},
  };
  const context = {
    chrome: {
      runtime: {
        lastError: null,
        onConnect,
        connectNative() { return nativePort; },
      },
      alarms: { create() {}, onAlarm: new EventHook() },
      tabs: { onActivated: new EventHook() },
      windows: { onFocusChanged: new EventHook() },
    },
    console: { log() {}, warn() {}, error() {} },
    setTimeout() {},
  };
  vm.createContext(context);
  const source = fs.readFileSync(path.join(__dirname, "..", "background.js"), "utf8");
  vm.runInContext(source, context, { filename: "background.js" });
  return { connect: (port) => onConnect.emit(port), reply: (message) => nativeMessages.emit(message) };
}

test("replacement reply returns only to the frame that supplied active text", () => {
  const background = loadBackground();
  const documentFrame = createContentPort(42, 0);
  const editorIframe = createContentPort(42, 7);
  const unrelatedFrame = createContentPort(42, 9);
  background.connect(documentFrame);
  background.connect(editorIframe);
  background.connect(unrelatedFrame);

  editorIframe.send({ type: "textUpdate", text: "Jeg liker piza." });
  background.reply({ action: "replace", tabId: 42, expected: "piza", text: "pipa" });

  assert.equal(documentFrame.sent.length, 0);
  assert.equal(unrelatedFrame.sent.length, 0);
  assert.deepEqual(editorIframe.sent, [{ action: "replace", tabId: 42, expected: "piza", text: "pipa" }]);
});

test("legacy replacement reply uses the latest active editor frame without broadcasting", () => {
  const background = loadBackground();
  const documentFrame = createContentPort(42, 0);
  const editorIframe = createContentPort(42, 7);
  background.connect(documentFrame);
  background.connect(editorIframe);

  editorIframe.send({ type: "textUpdate", text: "Jeg liker piza." });
  background.reply({ action: "replace", expected: "piza", text: "pipa" });

  assert.equal(documentFrame.sent.length, 0);
  assert.deepEqual(editorIframe.sent, [{ action: "replace", expected: "piza", text: "pipa" }]);
});

test("log messages from another tab cannot redirect a legacy replacement", () => {
  const background = loadBackground();
  const editor = createContentPort(42, 0);
  const otherTab = createContentPort(77, 0);
  background.connect(editor);
  background.connect(otherTab);

  editor.send({ type: "textUpdate", text: "Jeg liker piza." });
  otherTab.send({ type: "log", message: "background editor is idle" });
  background.reply({ action: "replace", expected: "piza", text: "pipa" });

  assert.deepEqual(editor.sent, [{ action: "replace", expected: "piza", text: "pipa" }]);
  assert.equal(otherTab.sent.length, 0);
});
