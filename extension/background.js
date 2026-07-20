// Spell background service worker — bridges content scripts to native messaging host

const NATIVE_HOST = "com.cognio.spell.bridge";
let nativePort = null;
let contentPorts = new Map(); // tabId -> [port, port, ...] (multiple frames per tab)
let lastActiveTabId = null;
// The desktop bridge can only act on one editor snapshot at a time. Keep the
// exact content port that supplied that snapshot so a reply returns to the
// same frame instead of every iframe in the tab.
let latestTextPortByTab = new Map();

// Send a message to every content port in a tab, pruning any that have
// become stale (e.g. their page was moved into Chrome's back/forward
// cache). Without pruning, contentPorts keeps growing and Chrome logs
//   "Unchecked runtime.lastError: The page keeping the extension port
//    is moved into back/forward cache, so the message channel is closed."
// for every postMessage to a bfcached page — observed 2026-05-18.
function sendToTabPorts(tabId, msg) {
  const ports = contentPorts.get(tabId);
  if (!ports) return;
  const alive = [];
  for (const p of ports) {
    try {
      p.postMessage(msg);
      alive.push(p);
    } catch (_) {
      // Port dead — don't keep it. Reading lastError marks the error as
      // "checked" so Chrome stops emitting the warning even if the
      // disconnect arrives after we already gave up on the port.
      void chrome.runtime.lastError;
    }
  }
  if (alive.length === 0) contentPorts.delete(tabId);
  else contentPorts.set(tabId, alive);
  disconnectNativeIfIdle();
}

function sendToLatestTextPort(tabId, msg) {
  const port = latestTextPortByTab.get(tabId);
  if (!port) return false;
  try {
    port.postMessage(msg);
    return true;
  } catch (_) {
    void chrome.runtime.lastError;
    latestTextPortByTab.delete(tabId);
    return false;
  }
}

function hasContentPorts() {
  for (const ports of contentPorts.values()) {
    if (ports.length > 0) return true;
  }
  return false;
}

function disconnectNativeIfIdle() {
  if (nativePort && !hasContentPorts()) {
    console.log("Spell: native host idle; disconnecting");
    try { nativePort.disconnect(); } catch (_) {}
    nativePort = null;
  }
}

function connectNative() {
  if (nativePort) return;
  try {
    nativePort = chrome.runtime.connectNative(NATIVE_HOST);
    console.log("Spell: native host connected");
    nativePort.onMessage.addListener((msg) => {
      // Forward replace actions to content script
      if (msg.action === "replace") {
        console.log("Spell replace:", JSON.stringify(msg));
        const targetTabId = Number.isInteger(msg.tabId) && msg.tabId > 0 ? msg.tabId : null;
        if (targetTabId) {
          if (!sendToLatestTextPort(targetTabId, msg)) {
            console.warn("Spell: replacement target frame is not connected:", targetTabId);
          }
        } else if (lastActiveTabId && sendToLatestTextPort(lastActiveTabId, msg)) {
          // Legacy replies have no tab ID. The latest text update is still
          // the only safe origin because broadcasting can replace text in
          // unrelated iframes and tabs.
        } else {
          console.warn("Spell: replacement has no active source frame");
        }
      }
    });
    nativePort.onDisconnect.addListener(() => {
      console.log("Spell native host disconnected:", chrome.runtime.lastError?.message);
      nativePort = null;
      // Reconnect only while a content script is attached. Keeping the
      // native host open with no page-side consumer leaves native_bridge.exe
      // running after the desktop app is uninstalled, which locks
      // AppData\Local\Spell\current and blocks the next install.
      if (hasContentPorts()) setTimeout(connectNative, 500);
    });
  } catch (e) {
    console.error("Spell: failed to connect native host:", e);
    nativePort = null;
  }
}

chrome.runtime.onConnect.addListener((port) => {
  if (port.name !== "spell-content") return;

  const tabId = port.sender?.tab?.id;
  if (tabId) {
    if (!contentPorts.has(tabId)) contentPorts.set(tabId, []);
    contentPorts.get(tabId).push(port);
  }

  port.onMessage.addListener((msg) => {
    if (msg.type === "textUpdate" && tabId) {
      lastActiveTabId = tabId;
      latestTextPortByTab.set(tabId, port);
    }
    connectNative();
    if (nativePort) {
      // Forward ALL messages (textUpdate + log) to native host
      if (msg.type === "textUpdate") msg.tabId = tabId;
      try {
        nativePort.postMessage(msg);
      } catch (e) {
        // postMessage on a half-dead native port silently fails — the port
        // object exists but the underlying pipe is closed (typical pattern
        // after a long idle period when Chrome's MV3 service worker
        // suspended then re-awakened, or after the user restarted the
        // desktop binary). Without clearing nativePort here, every later
        // postMessage hits the same dead pipe and the user sees "browser
        // session broken, even desktop restart doesn't help" — reported
        // 2026-05-15.
        //
        // Drop the reference so the next message goes through
        // connectNative() at line above and gets a fresh native_bridge.
        console.warn("Spell: nativePort postMessage failed, reconnecting:", e?.message || e);
        try { nativePort.disconnect(); } catch (_) {}
        nativePort = null;
        // Immediate reconnect attempt + buffer this message for resend
        connectNative();
        if (nativePort) {
          try { nativePort.postMessage(msg); } catch (_) {}
        }
      }
    }
  });

  port.onDisconnect.addListener(() => {
    // Drain runtime.lastError so Chrome doesn't emit
    //   "Unchecked runtime.lastError: The page keeping the extension
    //    port is moved into back/forward cache, ..."
    // for normal bfcache disconnects. The lastError property has
    // meaningful info ("Back/forward cache" vs "Receiving end does not
    // exist") that we don't act on but should at least read.
    void chrome.runtime.lastError;
    if (tabId && contentPorts.has(tabId)) {
      const ports = contentPorts.get(tabId).filter(p => p !== port);
      if (ports.length === 0) contentPorts.delete(tabId);
      else contentPorts.set(tabId, ports);
    }
    if (tabId && latestTextPortByTab.get(tabId) === port) {
      latestTextPortByTab.delete(tabId);
    }
    disconnectNativeIfIdle();
  });
});

// Keep the service worker awake enough to repair active page connections.
//
// MV3 service workers go to sleep after ~30 s of inactivity. The keepalive
// alarm wakes us up every 6 s and reconnects the native bridge only when a
// content script is currently attached. A page-side textUpdate still calls
// connectNative() immediately, so bridge recovery is driven by real input
// instead of a permanent background native_bridge.exe process.
chrome.alarms.create("keepalive", { periodInMinutes: 0.1 });
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "keepalive") {
    if (hasContentPorts()) connectNative();
    else disconnectNativeIfIdle();
  }
});

// Also wake the SW + bridge on tab activation / window focus so the
// "switch to a tab the bridge has lost" scenario heals quickly without
// waiting up to 6 s for the next alarm tick.
chrome.tabs.onActivated.addListener(() => {
  if (hasContentPorts()) connectNative();
});
chrome.windows.onFocusChanged.addListener(() => {
  if (hasContentPorts()) connectNative();
});
