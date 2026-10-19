# CLAUDE.md

Extracts text and embedded images from corrupt DOCX / XLSX / PPTX files
that Office cannot open. The recovery engine walks ZIP local file headers
directly, so it works even when the central directory is missing /
truncated and even when individual entries fail their CRC. Originally a
Windows-only Delphi 7 utility — **the TypeScript core under `core/` is now
the canonical implementation**; the Delphi source is preserved only for
historical reference.

## Repo map

- `core/` — TypeScript recovery engine. **Edits go here.** Powers every
  shipping platform.
- `src/` — additional shared source; `src/legacy-delphi/` is the
  read-only Delphi 7 source kept for reference (don't modify).
- `web/`, `web-app/` — web SPA + installable PWA bundle.
- `desktop/` — Electron wrapper (built with electron-builder). Produces
  Windows `.exe` + portable `.zip`, macOS `.dmg`, Linux `.AppImage` +
  `.deb`.
- `mobile/` — Capacitor wrapper. Produces unsigned Android `.apk` and an
  Xcode project for iOS.
- `releases/` — pre-packaged release archives committed to the repo.
- `scripts/` — release packaging helpers.
- `tests/` — recovery tests + `MANUAL.md` for iOS manual steps.
- `package.json` — root npm workspace; `npm run build:web`,
  `npm run build:android-debug -w mobile`, etc.
- `xpstyle.xml` — Windows visual styles manifest for the legacy Delphi
  binary.
- `.github/workflows/` — `build.yml` (CI), `pages.yml` (deploy SPA to
  Pages on push to `main`), `release.yml` (build per-platform installers
  on `v*` tag).

## Branch policy

Work on the assigned feature branch:

1. Commit and push the feature branch.
2. **Open a PR from the feature branch to `main`** using the GitHub MCP
   tools (`mcp__github__create_pull_request`). Do not merge directly —
   the maintainer reviews and merges.
3. CI runs on the PR; Pages and Release pipelines fire from `main` only.

## Releasing

- Push a `v*` tag to `main` (or use Actions → Release → workflow_dispatch)
  to build:
  - Web: hosted SPA + PWA bundle (via `npm run build:web`)
  - Desktop: `cd desktop && npm install && npx electron-builder --win/mac/linux …`
  - Mobile: `npm run build:android-debug -w mobile`; iOS is a manual Xcode build

## Verifying changes

- `npm test` at the repo root runs the recovery test suite.
- For SPA changes, `npm run build:web` then serve `web/` and exercise the
  three pipeline stages (standard ZIP → header scan → XML repair) on a
  known-corrupt fixture.
- Desktop / mobile wrappers should not contain recovery logic — if a fix
  requires changes outside `core/`, you're probably patching the wrapper
  when you should patch the core.

## Gotchas

- **Never edit `src/legacy-delphi/`** — it's historical reference. Any
  fix must go through the TypeScript core, which is what ships.
- The recovery engine must not throw on bad bytes. Returning partial data
  is correct; raising an exception breaks the "always salvage what you
  can" contract.
- iOS builds are *unsigned* and ship as an Xcode project — signing is the
  user's responsibility. Don't add hard-coded signing identifiers.
- Capacitor + Electron both consume the same `core/` build output. Keep
  the public API stable, or pin both wrappers to the new shape in the
  same PR.
