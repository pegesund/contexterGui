// Spell background service worker — bridges content scripts to native messaging host

const NATIVE_HOST = "com.cognio.spell.bridge";
let nativePort = null;
let contentPorts = new Map(); // tabId -> [port, port, ...] (multiple frames per tab)
let lastActiveTabId = null;

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
          for (const cPort of contentPorts.get(lastActiveTabId)) {
            try { cPort.postMessage(msg); } catch(e) {}
          }
        } else {
          for (const [tabId, ports] of contentPorts) {
            for (const cPort of ports) {
              try { cPort.postMessage(msg); } catch(e) {}
            }
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
      try { nativePort.postMessage(msg); } catch(e) {}
    }
  });

  port.onDisconnect.addListener(() => {
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
