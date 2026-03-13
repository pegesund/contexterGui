// NorskTale background service worker — bridges content scripts to native messaging host

const NATIVE_HOST = "com.norsktale.bridge";
let nativePort = null;
let contentPorts = new Map(); // tabId -> port
let lastActiveTabId = null;

function connectNative() {
  if (nativePort) return;
  try {
    nativePort = chrome.runtime.connectNative(NATIVE_HOST);
    nativePort.onMessage.addListener((msg) => {
      // Every response from native host could be a replace action or just an ack
      if (msg.action === "replace") {
        console.log("NorskTale replace:", JSON.stringify(msg), "tabs:", [...contentPorts.keys()]);
        // Send to the last active tab
        if (lastActiveTabId && contentPorts.has(lastActiveTabId)) {
          try {
            contentPorts.get(lastActiveTabId).postMessage(msg);
            console.log("Sent replace to tab", lastActiveTabId);
          } catch(e) {
            console.log("Failed to send to tab", lastActiveTabId, e);
          }
        } else {
          // Fallback: send to all
          for (const [tabId, cPort] of contentPorts) {
            try { cPort.postMessage(msg); } catch(e) {}
          }
        }
      }
    });
    nativePort.onDisconnect.addListener(() => {
      console.log("NorskTale native host disconnected:", chrome.runtime.lastError?.message);
      nativePort = null;
    });
  } catch (e) {
    console.error("NorskTale: failed to connect native host:", e);
    nativePort = null;
  }
}

chrome.runtime.onConnect.addListener((port) => {
  if (port.name !== "norsktale-content") return;

  const tabId = port.sender?.tab?.id;
  if (tabId) contentPorts.set(tabId, port);

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
    if (tabId) contentPorts.delete(tabId);
  });
});

// Keep service worker alive
chrome.alarms.create("keepalive", { periodInMinutes: 0.1 });
chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "keepalive" && contentPorts.size > 0) {
    connectNative();
  }
});
