# web-apps build overlay (v1 wrapper)

Three small patches to `web-apps/build/{editor}.json` that strip
~82 MB of disk footprint per editor without touching JS bundle output.
Applied via `apply-overlay.sh` against an the vendor/web-apps clone.

## Why an overlay (vs forking web-apps)

Per [10-web-apps-investigation.md §10.0.3](../../analysis/plan/phases/10-web-apps-investigation.md#1003-deliverable--measured-build-experiments-documenteditor),
all the bloat we want gone is removable through `build/{editor}.json`
edits — no source modifications. Keeping web-apps unmodified means
upstream `git pull` never conflicts with us; we re-apply the overlay
after each pull and rebuild.

## What the overlay changes

For each editor (currently: documenteditor; spreadsheet/presentation
in 10.1.2):

| Edit | Before | After | Saves |
|---|---|---|---|
| `tasks.deploy` | `["increment-build","deploy-app-main","deploy-app-mobile","deploy-app-forms","deploy-app-embed"]` | `["increment-build","deploy-app-main"]` | ~14 MB (mobile/embed/forms variants) |
| `main.copy.localization[0].src` | `"*"` | `"en.json"` | ~12 MB (44 unused locales) |
| `main.copy.help` | full glob over per-locale HTML help | 1-file stub (`en/Contents.json`) | ~78 MB (in-app help docs) |

JS bundle output (`app.js` + `code.js`) is byte-identical to baseline
— this overlay only changes which static assets get copied alongside.

## Usage

```bash
# from anywhere, defaults to the sibling web-apps clone
wrapper/v1/web-apps-overlay/apply-overlay.sh

# explicit path
wrapper/v1/web-apps-overlay/apply-overlay.sh \
  --web-apps /path/to/web-apps \
  --editors documenteditor

# revert (pristine build configs)
git -C /path/to/web-apps checkout HEAD -- build/*.json
```

## Per-release update flow

```bash
cd /path/to/web-apps
git pull                                                         # 1. pull upstream
git checkout HEAD -- build/*.json                                # 2. discard previous overlay
cd build && npm install                                          # 3. refresh deps if needed
cd /path/to/sdkjs                                                # 4. our repo
wrapper/v1/web-apps-overlay/apply-overlay.sh                      # 5. re-apply overlay
cd /path/to/web-apps/build && npx grunt deploy-documenteditor    # 6. rebuild
```

Target: < 5 minutes for a smooth-case re-apply.

## Caveats

- **`replace:prepareHelp` aborts on empty `copy.help[]`.** Hence the
  1-file stub. A cleaner upstream fix would patch `Gruntfile.js`'s
  `replace:prepareHelp` definition; we accept the few-KB cost
  instead.
- **`build` field in each JSON auto-increments on every Grunt run.**
  This shows up as a diff against HEAD even on no-op rebuilds. Don't
  commit it back upstream — `git checkout HEAD -- build/*.json`
  before each re-apply handles it.
- **Schema drift risk.** If OnlyOffice renames `tasks.deploy` or
  changes the `copy` structure, the jq filter silently produces
  bogus output. The script does not yet validate the post-patch
  shape; add a smoke check (e.g. `jq -e '.tasks.deploy | length == 2'`)
  if this becomes an issue.
