# Local cross-origin test

This directory has a tiny test page that simulates your messenger,
running on a **different origin** from the editor so the cross-origin
postMessage flow is exercised.

## Setup (one-time)

You need TWO servers running on different ports:

```
http://localhost:8080    ← editor (this repo's server/index.js, serves public/)
http://localhost:3000    ← test host (test/server.js, serves test/index.html)
```

Different ports = different origins in the browser's eyes, even on
the same hostname. That's enough to exercise cross-origin postMessage.

### 1. Build or stage `public/`

Either run the full build pipeline:

```bash
bash build/build.sh
```

Or copy files manually (see the "manual public/ setup" notes
script). Either way, you should end up with a `public/` directory
containing `edit.html`, `wrapper-*.js`, `x2t/`, `fonts/`, `sdkjs/`,
`web-apps/`, etc.

### 2. Bake the allowed host origin into `public/edit.html`

The editor's `edit.html` has an `__ALLOWED_HOST_ORIGIN__` placeholder
that the build pipeline normally substitutes. For local testing
against `http://localhost:3000`:

```bash
# macOS:
sed -i '' 's|__ALLOWED_HOST_ORIGIN__|http://localhost:3000|g' public/edit.html

# Linux:
sed -i 's|__ALLOWED_HOST_ORIGIN__|http://localhost:3000|g' public/edit.html
```

Verify:

```bash
grep -F "ALLOWED_HOST_ORIGIN = " public/edit.html
# Should show:  var ALLOWED_HOST_ORIGIN = 'http://localhost:3000';
```

If you see `'__ALLOWED_HOST_ORIGIN__'` still, the substitution didn't
happen — re-run the sed command.

### 3. Start both servers

In one terminal:

```bash
npm start             # editor at http://localhost:8080
```

In another:

```bash
npm run test:host     # test host at http://localhost:3000
```

### 4. Open the test page

```
http://localhost:3000/
```

You'll see:

```
Editor test host

Editor origin: http://localhost:8080  ·  Test host origin: http://localhost:3000

[Choose file] [Open in word] [...cell] [...slide]
[Save → download OOXML] [Close editor] | no file | no editor tab | clean

(log panel)
```

## Using the test page

1. **Pick a file** — `.docx`, `.xlsx`, or `.pptx` from your filesystem.
2. **Click an "Open" button** — opens a new tab at
   `http://localhost:8080/edit.html?type=word`. The editor loads,
   sends `ready`, and the test page replies with the file bytes via
   `postMessage`.
3. **Edit the doc** — typing should flip the dirty chip to "unsaved
   edits".
4. **Click "Save → download OOXML"** — sends `save-request`. The
   editor captures bytes, x2t reverse-converts to OOXML, posts back as
   `saved`. Your browser downloads the result.
5. **Re-open the downloaded file** — verify your edits persisted.
6. **Close the editor tab manually** — the test page detects within
   2 seconds and updates the chip. The editor's connection-lost modal
   should appear instantly when the messenger (this page) reloads or
   closes too.

## Different editor port?

If you run the editor on a different port, append `?editor=http://...`
to the test page URL:

```
http://localhost:3000/?editor=http://localhost:9090
```

…and update `__ALLOWED_HOST_ORIGIN__` in `public/edit.html`
accordingly.

## What's NOT being tested locally

- **Encryption.** The test page sends raw file bytes. Production
  flow has the messenger decrypt before send / encrypt before storing.
- **S3 storage.** The test page just downloads the saved OOXML.
- **Real authentication.** Anyone on the local network could open
  the test page. In production, only authenticated messenger users
  can trigger `window.open`.
- **TLS / HTTPS.** Local is plain HTTP. Production requires HTTPS
  for `window.opener` access in some browser policies (and for
  SharedArrayBuffer to work alongside COEP).

For all of those, exercise via the actual messenger build once
this smoke passes.

## Common issues

### "Cannot read properties of undefined" on editor load

Almost certainly a known cascade layer. Check the editor tab's
DevTools console for the specific error, then look it up in
the cascade layer documentation.

### Editor opens but immediately shows "open from messenger"

The editor sees `window.opener === null`. Causes:

- You opened `http://localhost:8080/edit.html` directly in the URL
  bar (correct behavior — refuse to mount).
- Pop-up blocker swallowed the `window.open` call. Check the test
  page's log: if it says `window.open returned null`, that's the
  cause. Allow pop-ups for `http://localhost:3000`.
- Some Chrome flag is treating `window.open` between same-domain
  different-port as cross-origin in a weird way. Try Firefox.

### Editor mounts but messages don't flow

Most common: `__ALLOWED_HOST_ORIGIN__` in `public/edit.html` doesn't
match `http://localhost:3000`. Check:

```bash
grep "ALLOWED_HOST_ORIGIN = " public/edit.html
```

It must read `'http://localhost:3000'` exactly (no trailing slash,
no path). If it says `'__ALLOWED_HOST_ORIGIN__'`, you forgot to run
the sed substitution.

### "Mixed Content" warnings

You're on `https://` test page but the editor is on `http://` (or
vice versa). Both must be HTTP for local dev or both HTTPS in
production.

### Connection-lost modal fires almost immediately

This used to happen if the editor was served with
`Cross-Origin-Opener-Policy: same-origin` AND opened cross-origin via
`window.open` — that combination severs `window.opener` and breaks
postMessage. The current `server/index.js` does NOT set COOP for
exactly this reason.

If you somehow re-introduced COOP `same-origin` (in nginx, a CDN, or
a proxy), the symptoms are:
- Modal appears within seconds of opening the editor
- Editor's DevTools console shows no `[edit] HOST_ORIGIN locked to ...`
  message OR shows it but no further activity
- Test host's log shows "ready" never arriving

Fix: remove COOP from the editor's response headers. HarfBuzz + x2t
WASM are single-threaded and don't need SharedArrayBuffer.

### Editor mounts but no pings arrive (heartbeat-timeout fires after 60s)

Origin mismatch. Check `public/edit.html`:

```bash
grep "ALLOWED_HOST_ORIGIN = " public/edit.html
# Must show: var ALLOWED_HOST_ORIGIN = 'http://localhost:3000';
# (the EXACT origin you're testing from, not the example value)
```

The editor's DevTools console will show a clear warning on the first
dropped ping — look for `[edit] DROPPING ping from ...`.

### Connection-lost modal fires immediately

The editor isn't receiving heartbeat `ping` messages. Check the
test page's log — if you see `← pong` lines, the heartbeat is
working both ways. If not, see the previous "messages don't flow"
section.
