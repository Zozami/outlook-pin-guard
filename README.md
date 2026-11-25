# Send PIN Guard

An Outlook add-in that blocks **sending** a message until a PIN/passphrase is entered.
The unlock is valid for **one send** and **60 seconds**.

- Works in **new Outlook for Windows**, **Outlook on the web**, and **classic Outlook for Windows** (current builds).
- Hosted from GitHub Pages: `https://zozami.github.io/outlook-pin-guard/`
- Manifest for sideloading: `manifest.xml`

> **Important:** this is a deterrent, not a security boundary. See
> [What this does and does not protect](#what-this-does-and-does-not-protect).

---

## How it works

1. You click **Send**.
2. Outlook fires the `OnMessageSend` event. The add-in checks the message's
   `sessionData` for a fresh unlock token.
3. If there is no valid token, the send is **blocked** and Outlook shows:
   **"Sending is locked."** with an **Enter PIN** button.
4. The button opens the task pane. You enter the PIN.
5. The PIN is checked against a salted PBKDF2-SHA256 hash in `config.js`
   (600,000 iterations, 16-byte salt).
6. On success, the task pane writes a timestamp to `sessionData` and sends the
   message. The send event consumes the timestamp, so one PIN entry = one send.

`SendMode="Block"` is intentional: if the add-in cannot load, **the send is
blocked** rather than silently allowed. `SoftBlock`/`PromptUser` would let the
message out when the add-in is unavailable.

---

## Setup (one time)

### 1. Choose the PIN

Use a **long passphrase (12+ characters)**, not a 4–6 digit PIN. The
`config.js` file is public on GitHub Pages, so a short PIN can be cracked
offline with a GPU. A long passphrase is what makes the public hash safe.

### 2. Generate `config.js`

`make-hash.js` is deliberately **not** in this repo. Keep it in the parent
folder (`..\make-hash.js`). From this folder:

```powershell
node ..\make-hash.js "YOUR LONG PASSPHRASE" --out config.js
```

Check it:

```powershell
Get-Content config.js
```

You should see a real 32-hex-character `saltHex` and a 64-hex-character
`hashHex`, not `PASTE_SALT`/`PASTE_HASH`.

> The PIN itself is never written to disk by the add-in and never committed.
> Only the salt and hash are published.

### 3. Publish to GitHub Pages

This repo is already set up for it. If you need to do it from scratch:

```powershell
git init
git add .
git commit -m "pin guard"
git branch -M main
gh repo create outlook-pin-guard --public --source=. --push
gh api -X POST repos/Zozami/outlook-pin-guard/pages `
  -f "source[branch]=main" -f "source[path]=/"
```

Then wait 1–3 minutes and verify every file returns HTTP 200:

```powershell
$base = "https://zozami.github.io/outlook-pin-guard"
"commands.html","launchevent.js","taskpane.html","taskpane.js","config.js",
"icon-16.png","icon-32.png","icon-64.png","icon-80.png","icon-128.png" |
  ForEach-Object {
    $code = (Invoke-WebRequest "$base/$_" -Method Head -UseBasicParsing).StatusCode
    "$_ $code"
  }
```

`manifest.xml` is **not** served from Pages; it is installed from the local file.

### 4. Install in Outlook

1. Sign in to your **company** Microsoft 365 account in a browser.
2. Open <https://aka.ms/olksideload>.
3. In **Add-Ins for Outlook**, go to **My add-ins** → **Custom add-ins**.
4. Choose **Add a custom add-in** → **Add from file**.
5. Select `manifest.xml` from this folder and accept the prompts.
6. Fully close and reopen Outlook. Add-in changes can take a while to appear.

If "Add a custom add-in" is missing or errors, your admin has disabled
sideloading/custom add-ins (Exchange role **My Custom Apps**). That is a tenant
setting; this add-in cannot override it.

### 5. Test

1. Compose a message to yourself and click **Send**.
   Expected: **"Sending is locked."** and an **Enter PIN** button.
2. Click **Enter PIN**, enter a **wrong** PIN.
   Expected: **"Wrong PIN."** and nothing is sent.
3. Enter the **correct** PIN.
   Expected: **"Sent."** and the message is delivered.
4. Compose another message and click **Send** twice without entering the PIN.
   Expected: blocked both times.

If the message simply sends, the add-in is not activating. Check:
- The add-in is installed and enabled (File → Info → Manage Add-ins).
- You are using new Outlook or Outlook on the web (mobile is not supported).
- The Pages URL loads in the same browser session (`config.js`, `taskpane.js`).
- Your tenant allows custom add-ins.

---

## Updating

### Change the PIN

```powershell
node ..\make-hash.js "NEW LONG PASSPHRASE" --out config.js
git add config.js
git commit -m "update PIN hash"
git push
```

If the add-in does not pick up the new file, remove and re-add it in
**My add-ins**, or clear the Office cache. Bump `<Version>` in `manifest.xml`
when you change the manifest itself (for example, the hosted URLs).

### Change the hosted files

GitHub Pages serves the new files after a push. Outlook caches add-in files
aggressively; remove/re-add the add-in if the task pane looks stale.

---

## Troubleshooting

| Symptom | Likely cause |
| --- | --- |
| "Add a custom add-in" is missing | Tenant blocks custom add-ins (`My Custom Apps` role). Ask IT. |
| Add-in installs but Send is never blocked | Event-based activation did not load. Check the add-in is enabled and the Pages URL is reachable from the Outlook machine. In classic Outlook, check Event Viewer → Windows Logs → Application, Event ID 63. |
| "PIN is not configured in config.js yet" | `config.js` still has `PASTE_SALT`/`PASTE_HASH`, or the file did not deploy. |
| "Wrong PIN." for the correct PIN | `config.js` on Pages does not match the PIN you typed. Regenerate and redeploy. |
| Task pane is blank | `office.js` blocked, or the Pages site is blocked by the corporate proxy/firewall. |
| Send is blocked even after a correct PIN | The unlock is 60 seconds / one send. Re-enter the PIN. If it persists, check that the task pane and the send handler are the same add-in (same `<Id>`). |

### Tenant policy for offline/blocked sends

For Outlook on the web and new Outlook on Windows, an admin can set
`OnSendAddinsEnabled` on the OWA mailbox policy so installed add-ins must run
on outgoing mail, even offline. For classic Outlook, `SendMode="Block"` already
fails closed. These are tenant-level settings; a user cannot set them.

---

## What this does and does not protect

**What it does**

- Blocks a normal Send in Outlook until the correct PIN is entered.
- Requires the PIN on **every** send (one unlock = one send).
- Fails closed if the add-in cannot load (`SendMode="Block"`).
- Stops a casual person at your unlocked PC from sending mail from your open
  Outlook window without the PIN.

**What it does not do**

- **You can remove or disable the add-in.** A user-installed add-in is under
  your account's control. Anyone who can use Outlook can open
  **Get Add-ins → My add-ins**, remove PIN Guard, and send. Only an
  admin-deployed add-in with users unable to remove it makes this stronger.
- **Outlook mobile does not support Smart Alerts.** Sending from the phone
  bypasses this entirely.
- **Other mail clients and APIs bypass it.** IMAP/SMTP, PowerShell, Graph,
  Thunderbird, etc. are not covered.
- **The PIN is verified in JavaScript.** Someone with devtools and knowledge of
  the app can write the unlock token without the PIN. The hash also lives on a
  public GitHub Pages site, so a short PIN can be cracked offline.
- **Send Later / scheduled send** is not converted to an immediate send, but it
  is still only checked when the send event fires.
- **Meeting invites are not covered** (the manifest only handles messages).

For real protection, use:
- a locked Windows session (Win+L) and a short idle lock timeout;
- an admin-deployed (mandatory) add-in for everyone;
- Microsoft Purview DLP / mail flow rules for outbound mail;
- Conditional Access and device compliance for mailbox access.

---

## Files

| File | Purpose |
| --- | --- |
| `manifest.xml` | Add-in manifest (install this file in Outlook). |
| `commands.html` | Runtime page for the send event. |
| `launchevent.js` | `OnMessageSend` handler. Blocks unless a fresh unlock exists. |
| `taskpane.html` / `taskpane.js` | PIN entry and unlock/send UI. |
| `config.js` | Salt + PBKDF2 hash + iterations. Generated by `make-hash.js`. |
| `icon-*.png` | Add-in icons. |

`make-hash.js` is intentionally excluded from the repo and must never be
committed. The plain-text passphrase must never be committed either.

---

## License

Private use. No warranty.
