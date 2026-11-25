// Send PIN Guard - task pane
// Verifies the PIN locally (PBKDF2 hash), then unlocks this one message for 60 seconds.
const UNLOCK_KEY = "pinGuardUnlockedAt";
const MAX_PIN_LENGTH = 256;
const DELAY_AFTER_FAILURES = 3;
const DELAY_MS = 30000;

let busy = false;
let failures = 0;

const hex = (buf) =>
  Array.from(new Uint8Array(buf))
    .map((b) => b.toString(16).padStart(2, "0"))
    .join("");

const fromHex = (h) => new Uint8Array(h.match(/.{2}/g).map((x) => parseInt(x, 16)));

// Compares two hex strings without leaking timing information.
function sameHex(a, b) {
  if (typeof a !== "string" || typeof b !== "string" || a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

function configIsReady() {
  return (
    typeof PIN_CONFIG !== "undefined" &&
    typeof PIN_CONFIG.saltHex === "string" &&
    /^[0-9a-f]{32}$/i.test(PIN_CONFIG.saltHex) &&
    typeof PIN_CONFIG.hashHex === "string" &&
    /^[0-9a-f]{64}$/i.test(PIN_CONFIG.hashHex) &&
    Number.isInteger(PIN_CONFIG.iterations) &&
    PIN_CONFIG.iterations > 0
  );
}

async function pinIsValid(pin) {
  const key = await crypto.subtle.importKey(
    "raw",
    new TextEncoder().encode(pin),
    "PBKDF2",
    false,
    ["deriveBits"]
  );
  const bits = await crypto.subtle.deriveBits(
    {
      name: "PBKDF2",
      salt: fromHex(PIN_CONFIG.saltHex),
      iterations: PIN_CONFIG.iterations,
      hash: "SHA-256",
    },
    key,
    256
  );
  return sameHex(hex(bits), PIN_CONFIG.hashHex);
}

const msgEl = () => document.getElementById("msg");

function setMsg(text, kind) {
  const el = msgEl();
  el.textContent = text;
  el.className = kind || "";
}

function setBusy(value) {
  busy = value;
  const input = document.getElementById("pin");
  const btn = document.getElementById("go");
  input.disabled = value;
  btn.disabled = value;
  btn.textContent = value ? "Working…" : "Unlock and send";
}

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function sessionData() {
  const item = Office.context.mailbox.item;
  if (!item || !item.sessionData) throw new Error("This Outlook client doesn't support session data.");
  return item.sessionData;
}

function setUnlock() {
  return new Promise((resolve, reject) => {
    sessionData().setAsync(UNLOCK_KEY, String(Date.now()), (r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(r.error && r.error.message ? r.error.message : "Could not unlock this message."));
    });
  });
}

function clearUnlock() {
  try {
    sessionData().removeAsync(UNLOCK_KEY, () => {});
  } catch (_) {
    // Best effort: the token also expires on its own after 60 seconds.
  }
}

function hasDelayedDelivery() {
  return new Promise((resolve) => {
    let delay;
    try {
      delay = Office.context.mailbox.item.delayDeliveryTime;
    } catch (_) {
      resolve(false);
      return;
    }
    if (!delay || typeof delay.getAsync !== "function") {
      resolve(false);
      return;
    }
    try {
      delay.getAsync((r) => {
        resolve(r.status === Office.AsyncResultStatus.Succeeded && !!r.value);
      });
    } catch (_) {
      resolve(false);
    }
  });
}

function sendItem() {
  const item = Office.context.mailbox.item;
  if (typeof item.sendAsync !== "function") {
    // Older clients: leave the unlock in place and let the user click Send.
    return Promise.resolve(false);
  }
  return new Promise((resolve, reject) => {
    item.sendAsync((r) => {
      if (r.status === Office.AsyncResultStatus.Succeeded) resolve(true);
      else reject(new Error(r.error && r.error.message ? r.error.message : "Outlook could not send the message."));
    });
  });
}

async function unlock() {
  if (busy) return;

  const input = document.getElementById("pin");
  const pin = input.value;

  if (!configIsReady()) {
    setMsg("PIN is not configured in config.js yet. The add-in owner needs to finish setup.", "err");
    return;
  }
  if (!pin) {
    setMsg("Enter your PIN.", "err");
    input.focus();
    return;
  }
  if (pin.length > MAX_PIN_LENGTH) {
    setMsg("Wrong PIN.", "err");
    input.value = "";
    return;
  }

  setBusy(true);
  setMsg("Checking…", "info");

  try {
    if (failures >= DELAY_AFTER_FAILURES) {
      setMsg("Too many attempts. Wait 30 seconds.", "err");
      await delay(DELAY_MS);
      failures = 0;
    }

    if (!(await pinIsValid(pin))) {
      failures++;
      input.value = "";
      setMsg(
        failures >= DELAY_AFTER_FAILURES
          ? "Wrong PIN. The next attempt will be delayed."
          : "Wrong PIN.",
        "err"
      );
      return;
    }

    failures = 0;
    input.value = "";
    setMsg("Unlocked. Sending…", "info");

    await setUnlock();
    // Give the launch-event runtime a moment to see the sessionData write.
    await delay(400);

    // Don't turn a "Send later" into an immediate send.
    if (await hasDelayedDelivery()) {
      setMsg("Unlocked. Click Send within 60 seconds to keep the scheduled delivery time.", "ok");
      return;
    }

    let sent = true;
    try {
      sent = await sendItem();
    } catch (e) {
      // Don't leave a valid unlock behind if the send failed.
      clearUnlock();
      throw e;
    }

    if (sent) setMsg("Sent.", "ok");
    else setMsg("Unlocked. Click Send within 60 seconds.", "ok");
  } catch (e) {
    clearUnlock();
    setMsg("Error: " + (e && e.message ? e.message : String(e)), "err");
  } finally {
    setBusy(false);
  }
}

function init() {
  const input = document.getElementById("pin");
  const btn = document.getElementById("go");
  btn.addEventListener("click", unlock);
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      unlock();
    }
  });
  input.focus();
  if (!configIsReady()) {
    setMsg("PIN is not configured in config.js yet. The add-in owner needs to finish setup.", "err");
  }
}

if (typeof Office !== "undefined") {
  Office.onReady(init);
} else {
  document.addEventListener("DOMContentLoaded", () =>
    setMsg("Outlook did not load. Open this pane from the Send PIN Guard add-in.", "err")
  );
}
