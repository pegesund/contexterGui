// Spell background service worker — bridges content scripts to native messaging host

const NATIVE_HOST = "com.cognio.spell.bridge";
let nativePort = null;
let contentPorts = new Map(); // tabId -> [port, port, ...] (multiple frames per tab)
let lastActiveTabId = null;

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
        if (lastActiveTabId && contentPorts.has(lastActiveTabId)) {
          sendToTabPorts(lastActiveTabId, msg);
        } else {
          for (const tabId of [...contentPorts.keys()]) {
            sendToTabPorts(tabId, msg);
          }
        }
      }
    });
    nativePort.onDisconnect.addListener(() => {
      console.log("Spell native host disconnected:", chrome.runtime.lastError?.message);
      nativePort = null;
      // Always retry — the user might still be in a writing app; switching
      // tabs / navigating breaks content-script ports briefly and used to
      // leave the native bridge dead until a new tab loaded a fresh script.
      // Previously this only retried when contentPorts.size > 0, but
      // service-worker sleep cycles cause genuine disconnects from any
      // state, and the bridge dying mid-session means typing in a tab
      // that ALREADY has a content script silently fails to reach the
      // desktop (extension SW handles the postMessage but native port
      // is dead → message dropped).
      setTimeout(connectNative, 500);
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
    lastActiveTabId = tabId;
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
  });
});

// Keep service worker + native bridge alive.
//
// MV3 service workers go to sleep after ~30 s of inactivity. The keepalive
// alarm wakes us up every 6 s, which:
//   (1) postpones the SW sleep deadline (because the SW had work to do)
//   (2) lets us reconnect the native bridge if it died
//
// Previously the alarm only called connectNative() when contentPorts.size > 0.
// That broke the common case where the user briefly switches away from a
// tab with a Spell content script: the script's port detaches, SW sleeps,
// native bridge gets EOF on stdin and exits, then when the user switches
// back and types, the SW wakes from the postMessage but the bridge is
// dead and the text gets dropped before reaching the desktop. Reported
// during 2026-05-15 testing: "caught errors on launch then stopped".
//
// Now we always try to keep the native bridge connected. The bridge is
// cheap (a tiny stdio process) and Chrome only spawns it if its
// connectNative call succeeds — so this is bounded.
chrome.alarms.create("keepalive", { periodInMinutes: 0.1 });
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "keepalive") {
    connectNative();
  }
});

// Also wake the SW + bridge on tab activation / window focus so the
// "switch to a tab the bridge has lost" scenario heals quickly without
// waiting up to 6 s for the next alarm tick.
chrome.tabs.onActivated.addListener(() => {
  connectNative();
});
chrome.windows.onFocusChanged.addListener(() => {
  connectNative();
});
