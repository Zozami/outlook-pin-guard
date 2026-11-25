// Runs on every Send. Fails CLOSED: anything other than a fresh, valid unlock blocks the send.
// Keep this file ES2016-safe: classic Outlook on Windows runs it in a JavaScript-only runtime.
const UNLOCK_KEY = "pinGuardUnlockedAt";
const UNLOCK_WINDOW_MS = 60 * 1000; // unlock is valid for 60 seconds and for ONE send

function blockSend(event) {
  event.completed({
    allowEvent: false,
    errorMessage: "Sending is locked. Select Enter PIN, then enter your PIN to send this message.",
    errorMessageMarkdown:
      "**Sending is locked.** Select **Enter PIN** and enter your PIN to send this message.",
    cancelLabel: "Enter PIN",
    commandId: "msgComposeOpenPaneButton"
  });
}

function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;
    if (!item || !item.sessionData) {
      blockSend(event);
      return;
    }

    item.sessionData.getAsync(UNLOCK_KEY, function (r) {
      try {
        if (r.status !== Office.AsyncResultStatus.Succeeded || !r.value) {
          blockSend(event);
          return;
        }

        const unlockedAt = Number(r.value);
        const age = Date.now() - unlockedAt;
        if (!isFinite(unlockedAt) || !(age >= 0 && age <= UNLOCK_WINDOW_MS)) {
          blockSend(event);
          return;
        }

        // Consume the unlock, then allow this one send.
        item.sessionData.removeAsync(UNLOCK_KEY, function () {
          event.completed({ allowEvent: true });
        });
      } catch (e) {
        blockSend(event);
      }
    });
  } catch (e) {
    blockSend(event);
  }
}

// Maps the function name in manifest.xml to this handler.
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
