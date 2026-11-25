// Runs on every Send. Fails CLOSED: anything other than a fresh, valid unlock blocks the send.
// Keep this file ES2016-safe: classic Outlook on Windows runs it in a JavaScript-only runtime.
const UNLOCK_KEY = "pinGuardUnlockedAt";
const UNLOCK_WINDOW_MS = 60 * 1000; // unlock is valid for 60 seconds and for ONE send
const COMPLETE_TIMEOUT_MS = 2000;   // never leave Outlook waiting on a missing key

function blockOptions() {
  return {
    allowEvent: false,
    errorMessage: "Sending is locked. Select Enter PIN, then enter your PIN to send this message.",
    errorMessageMarkdown:
      "**Sending is locked.** Select **Enter PIN** and enter your PIN to send this message.",
    cancelLabel: "Enter PIN",
    commandId: "msgComposeOpenPaneButton"
  };
}

function onMessageSendHandler(event) {
  let finished = false;
  let timer = null;

  function finish(allow) {
    if (finished) return;
    finished = true;
    if (timer) clearTimeout(timer);
    event.completed(allow ? { allowEvent: true } : blockOptions());
  }

  // The sessionData read can fail to call back when the key is absent.
  // Complete as "blocked" rather than letting Outlook hang on "processing".
  timer = setTimeout(function () { finish(false); }, COMPLETE_TIMEOUT_MS);

  try {
    const item = Office.context.mailbox.item;
    if (!item || !item.sessionData) {
      finish(false);
      return;
    }

    item.sessionData.getAsync(UNLOCK_KEY, function (r) {
      try {
        if (r.status !== Office.AsyncResultStatus.Succeeded || !r.value) {
          finish(false);
          return;
        }

        const unlockedAt = Number(r.value);
        const age = Date.now() - unlockedAt;
        if (!isFinite(unlockedAt) || !(age >= 0 && age <= UNLOCK_WINDOW_MS)) {
          finish(false);
          return;
        }

        // Consume the unlock, then allow this one send.
        item.sessionData.removeAsync(UNLOCK_KEY, function () {
          finish(true);
        });
      } catch (e) {
        finish(false);
      }
    });
  } catch (e) {
    finish(false);
  }
}

// Maps the function name in manifest.xml to this handler.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
