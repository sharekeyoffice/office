# `upstream-pins/` — the version knobs

These three JSON files are the **only** place the OnlyOffice version is
declared. The build clones/fetches each upstream input at the pinned ref; the
wrapper itself (`../wrapper/`) is in-repo, so a version bump is essentially
"edit a ref here + rebuild + re-verify the cascade".

| File | Pins | Consumed by |
|---|---|---|
| `sdkjs.json` | upstream **`ONLYOFFICE/sdkjs`** (document model + canvas SDK) | `build/01-pull-upstream.sh` → `git clone`/checkout into `vendor/sdkjs` |
| `web-apps.json` | upstream **`ONLYOFFICE/web-apps`** (editor UI shell) | `build/01-pull-upstream.sh` → `vendor/web-apps` (then grunt in `03`) |
| `x2t.json` | **CryptPad `onlyoffice-x2t-wasm`** release (the WASM converter) | provenance for `wrapper/v1/x2t/` (committed); `01` re-downloads from `url` only if the committed WASM is missing |

## Fields

- **`remote`** / **`ref`** (sdkjs, web-apps) — git remote + the exact ref to
  check out. `ref` should be a **commit SHA or your own tag**, not a moving
  branch/tag: AGPL §13 corresponding-source must correspond to the *exact*
  bytes served, and a pinned commit guarantees reproducibility (the build is
  byte-reproducible from a clean clone — see `../wrapper/REGENERATE.md`).
- **`url`** / **`sha512_url`** / **`version`** (x2t) — the CryptPad release zip
  + its checksum. AGPL-3.0; see `../THIRD-PARTY.md`.
- **`comment`** — free-text notes (not read by the build).

## How a bump works

1. Edit the `ref` in `sdkjs.json` + `web-apps.json` (and `x2t.json` if x2t
   moved). **Match `sdkjs`↔`web-apps` on major.minor**; match the **x2t major**
   to the sdkjs major.
2. Rebuild + re-verify the cascade — full procedure in
   [`../wrapper/REGENERATE.md`](../wrapper/REGENERATE.md).

> ⚠️ **Before bumping to OnlyOffice 9.4+:** that release adds AGPL §7
> attribution terms + a Trademark Policy. Review `../NOTICE` /
> `../THIRD-PARTY.md` and get legal sign-off on the branding posture.

## Current pins

- `sdkjs` — `ONLYOFFICE/sdkjs` @ `b2f0aa1d5c` (≈ v9.3.1.11+12) — the commit the
  wrapper is verified against.
- `web-apps` — `ONLYOFFICE/web-apps` @ `1c8ca998` (v9.3.1.1).
- `x2t` — CryptPad `v9.3.0+0`.
