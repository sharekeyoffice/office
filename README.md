# sharekey-office

Static-site deployment of the OnlyOffice v1 editor wrapper, hosted at
`office.hostname.com`. Serves the editor + bundled SDK + WASM
artifacts behind nginx in a Kubernetes pod.

The main app at `app.hostname.com` opens this site in a new
browser tab via `window.open(...)` and exchanges file bytes / save
events over cross-origin `postMessage`.

## Architecture (one paragraph)

`app.hostname.com` (proprietary app) → `window.open(...)` →
`office.hostname.com/edit.html?type=word&host=https://app.hostname.com`
→ this site loads the editor → cross-origin postMessage exchanges
file bytes (decrypted by main app before sending) ←→ saved bytes
(main app encrypts before storing in S3). No DocumentServer.

Full architecture and deployment notes are maintained separately.

## 60-second quick start

```bash
# 1. Configure the allowed main app origin (one-time per environment)
cp config.example.json config.json
# edit config.json — set "allowedHostOrigin" to your main app URL,
# e.g. "https://app.example.com"
# OR pass via env: export ALLOWED_HOST_ORIGIN=https://app.example.com

# 2. Pin upstream versions
cat upstream-pins/sdkjs.json     # adjust the ref
cat upstream-pins/web-apps.json

# 3. Build (~3 min: pulls upstream, runs grunt + bundle.js, prunes)
bash build/build.sh

# 4. Local dev server
npm start                         # serves dist/ on :8080

# 5. Open in a browser via the main app flow:
#    From your main app at the configured origin, run:
#    window.open('http://localhost:8080/edit.html?type=word', 'editor')
#    (no host parameter — the editor knows its allowed origin from build config)
```

## Production deploy (kubectl)

```bash
# Build + tag image
TAG=v9.3.0+0
docker build -t registry.hostname.com/sharekey-office:$TAG -f docker/Dockerfile .
docker push registry.hostname.com/sharekey-office:$TAG

# Then deploy the image with your own k8s manifests / orchestration.
```

Deployment manifests are environment-specific and not included in this repo.

## Layout

```
.
├── README.md                  this file
├── package.json               npm scripts (build, start)
├── server/
│   └── index.js               dev static server (Node stdlib)
├── build/                     build orchestration
│   ├── build.sh
│   ├── 01-pull-upstream.sh
│   ├── 02-build-bundles.sh
│   ├── 03-deploy-web-apps.sh
│   ├── 04-assemble-dist.sh
│   └── 05-prune.sh
├── config.example.json        per-environment config (copy → config.json, gitignored)
├── upstream-pins/             pinned versions (sdkjs/web-apps git refs, x2t release URL)
├── overlay/                   files copied LAST over dist/
│   ├── edit.html              customized: host-origin pinning + connection-lost modal
│   ├── connection-lost.css
│   ├── conflict-modal.css     concurrent-edit conflict UI
│   └── icons/                 per-file-type tab favicons (word/cell/slide .svg)
├── dist/                      build output — what nginx serves
└── docker/
    ├── Dockerfile             nginx:alpine
    └── nginx.conf             COOP/COEP, gzip_static, no SPA fallback
```

## Per-environment config

The allowed opener app origin (e.g. `https://app.example.com`) is
**baked into `dist/edit.html` at build time** — there's no runtime
URL parameter. This means:

- The editor only ever talks to the configured origin
- A malicious page that opens this editor cannot redirect it to its
  own origin (postMessages are silently dropped)
- Different environments (staging, prod) need different builds

To configure:

```bash
# Option A: config.json (gitignored)
cp config.example.json config.json
# edit "allowedHostOrigin" → "https://app.example.com"

# Option B: env var (overrides config.json — useful in CI)
export ALLOWED_HOST_ORIGIN=https://app.example.com
```

The build pipeline reads either, validates the shape
(`https?://host[:port]`), and substitutes the placeholder in
`overlay/edit.html`. If neither is set, the build fails with a clear
error.

## Updating to a new OnlyOffice release

1. Edit `upstream-pins/sdkjs.json` and `upstream-pins/web-apps.json`
   to reference the new tag.
2. Run `bash build/build.sh`.
3. Smoke test locally with `npm start`.
4. Tag a new image (`v9.4.0+0` or similar), push to registry.
5. Deploy the new image with your own k8s manifests / orchestration.

Detailed cascade verification: see
[wrapper/REGENERATE.md](wrapper/REGENERATE.md).

## What this repo is NOT

- It does NOT implement collab, encryption, key management, or
  storage.
- It does NOT contain auth or user identity. The editor is anonymous —
  it just renders bytes the main app sends it.
- It does NOT modify wrapper sources. Those live in the AGPL sdkjs
  fork; this repo just packages them.

## License & compliance

This deployment **bundles and serves AGPL-3.0 OnlyOffice code**, so the
whole repository is distributed under **GNU AGPL-3.0** — see
[`LICENSE`](LICENSE), [`NOTICE`](NOTICE) (ONLYOFFICE attribution +
statement of modifications), and [`THIRD-PARTY.md`](THIRD-PARTY.md)
(x2t, fonts, icons).

Our own tooling (build scripts, nginx config, Dockerfile,
`server/index.js`, overlay HTML/CSS, the wrapper glue under `wrapper/`)
is original work; because it is combined with and serves the AGPL'd
sdkjs/web-apps, it ships under AGPL-3.0 alongside them.

**This is a *modified* OnlyOffice redeployment** — not the original
software and not an official ONLYOFFICE product. ONLYOFFICE Docs **9.4+**
attaches AGPL §7 *additional terms* (prominent in-UI ONLYOFFICE credit,
stated modifications, license link) on top of AGPL; the separate
**Trademark Policy** governs the ONLYOFFICE name/logo. Obligations:

- **§13 (network use):** the served site must offer *corresponding
  source* to its users — link to the pinned upstream refs + this repo +
  `wrapper/sdkjs-patches/` + `build/` (an in-editor "Source code" link or
  `/source` page). Reproduce via [`wrapper/REGENERATE.md`](wrapper/REGENERATE.md).
- **§7 (9.4+):** show "Powered by ONLYOFFICE (modified by Sharekey)" +
  "ONLYOFFICE is a trademark of Ascensio System SIA" + a license link,
  baked into `dist/edit.html` via `overlay/edit.html`.
- **Trademark:** brand is "Sharekey", never "ONLYOFFICE"; no ONLYOFFICE
  logo/trade-dress as our branding; nominative use only.
- **Notices:** keep upstream license headers — but note `bundle.js` uses
  `legalComments:'none'` and currently strips them; see the caveat in
  `wrapper/REGENERATE.md`.
