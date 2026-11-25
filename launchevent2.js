// Runs on every Send. Fails CLOSED: anything other than a fresh, valid unlock blocks the send.
// Keep this file ES2016-safe: classic Outlook on Windows runs it in a JavaScript-only runtime.
const UNLOCK_KEY = "pinGuardUnlockedAt";
const UNLOCK_WINDOW_MS = 60 * 1000; // unlock is valid for 60 seconds and for ONE send
const COMPLETE_TIMEOUT_MS = 2500;   // never leave Outlook waiting on a missing key

const LOG_KEY = "pinGuardBlockedLog";
const LOG_MAX = 30;
const LOG_DEADLINE_MS = 800;

function blockOptions() {
  return {
    allowEvent: false,
    errorMessage: "Sending is locked. Select Enter PIN, then enter your PIN to send this message.",
    cancelLabel: "Enter PIN",
    commandId: "msgComposeOpenPaneButton"
  };
}

function formatRecipients(list) {
  if (!list || !list.length) return "";
  var out = [];
  for (var i = 0; i < list.length; i++) {
    var r = list[i];
    out.push((r.displayName ? r.displayName + " " : "") + "<" + r.emailAddress + ">");
  }
  return out.join("; ");
}

// Best-effort local audit entry. It never blocks the send decision for longer than LOG_DEADLINE_MS.
function recordBlocked(item, done) {
  var entry = { at: new Date().toISOString(), subject: "", to: "", cc: "", body: "" };
  var pending = 3;
  var finished = false;
  var timer = setTimeout(finish, LOG_DEADLINE_MS);

  function finish() {
    if (finished) return;
    finished = true;
    clearTimeout(timer);
    try {
      if (typeof localStorage !== "undefined") {
        var raw = localStorage.getItem(LOG_KEY);
        var list = raw ? JSON.parse(raw) : [];
        if (!list || typeof list.length !== "number") list = [];
        list.unshift(entry);
        localStorage.setItem(LOG_KEY, JSON.stringify(list.slice(0, LOG_MAX)));
      }
    } catch (e) {
      // Logging is optional; never break the send guard.
    }
    done();
  }

  function one() {
    pending--;
    if (pending <= 0) finish();
  }

  try {
    entry.subject = item.subject || "";
  } catch (e) {
    entry.subject = "";
  }

  try {
    item.to.getAsync(function (r) {
      try {
        if (r.status === Office.AsyncResultStatus.Succeeded) entry.to = formatRecipients(r.value);
      } catch (e) {}
      one();
    });
  } catch (e) {
    one();
  }

  try {
    item.cc.getAsync(function (r) {
      try {
        if (r.status === Office.AsyncResultStatus.Succeeded) entry.cc = formatRecipients(r.value);
      } catch (e) {}
      one();
    });
  } catch (e) {
    one();
  }

  try {
    item.body.getAsync("text", function (r) {
      try {
        if (r.status === Office.AsyncResultStatus.Succeeded) {
          entry.body = String(r.value || "").replace(/\s+/g, " ").slice(0, 500);
        }
      } catch (e) {}
      one();
    });
  } catch (e) {
    one();
  }
}

function onMessageSendHandler(event) {
  var finished = false;
  var loggingStarted = false;
  var timer = null;

  function complete(allow) {
    if (finished) return;
    finished = true;
    if (timer) clearTimeout(timer);

    var options = allow ? { allowEvent: true } : blockOptions();
    try {
      event.completed(options);
    } catch (e) {
      // If the client rejects the rich dialog options, at least complete with a plain block.
      try {
        event.completed({
          allowEvent: false,
          errorMessage: "Sending is locked. Open the Send PIN Guard pane to enter your PIN."
        });
      } catch (e2) {
        // Nothing else we can do; the client keeps the message unsent.
      }
    }
  }

  function block() {
    if (finished || loggingStarted) return;
    loggingStarted = true;
    try {
      var item = Office.context.mailbox.item;
      if (item) {
        recordBlocked(item, function () { complete(false); });
        return;
      }
    } catch (e) {}
    complete(false);
  }

  // Fail closed if anything below never calls back.
  timer = setTimeout(function () { complete(false); }, COMPLETE_TIMEOUT_MS);

  try {
    var item = Office.context.mailbox.item;
    if (!item || !item.sessionData) {
      block();
      return;
    }

    item.sessionData.getAllAsync(function (r) {
      try {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          block();
          return;
        }

        var data = r.value || {};
        var value = data[UNLOCK_KEY];
        if (!value) {
          block();
          return;
        }

        var unlockedAt = Number(value);
        var age = Date.now() - unlockedAt;
        if (!isFinite(unlockedAt) || !(age >= 0 && age <= UNLOCK_WINDOW_MS)) {
          block();
          return;
        }

        // Consume the unlock, then allow this one send.
        item.sessionData.removeAsync(UNLOCK_KEY, function () {
          complete(true);
        });
      } catch (e) {
        block();
      }
    });
  } catch (e) {
    block();
  }
}

// Maps the function name in manifest.xml to this handler.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
