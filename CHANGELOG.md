# Changelog

All notable changes to this project will be documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/);
versions follow [SemVer](https://semver.org/).

## [2.0.0] — 2026-05-09

Cross-platform modernization. The original Windows-only Delphi 7 app is
preserved untouched under `src/legacy-delphi/`; everything else is new.

### Added
- **Portable TypeScript core** (`core/`) that re-implements the corrupt-zip
  scanning + OOXML text/image extraction in a way that runs in Node, the
  browser, React Native / Capacitor, and Electron with no native bindings.
- **Static SPA** (`web/`) — drag-and-drop UI, runs entirely client-side, ships
  with a service worker + web app manifest so it's installable as a PWA on
  ChromeOS, Android, and desktops.
- **Electron desktop wrapper** (`desktop/`) for Windows (`.exe` + portable
  `.zip`), macOS (`.dmg`, unsigned), and Linux (`.AppImage` + `.deb`).
  Linux package also installs into ChromeOS Crostini.
- **Capacitor mobile wrapper** (`mobile/`) producing an Android debug APK and
  an iOS Xcode project (signed `.ipa` is documented but out of CI scope).
- **Corrupt fixture generator** (`scripts/make-fixtures.mjs`) producing four
  deliberately-broken DOCX/XLSX/PPTX files, each carrying a recoverable
  `HELLO_FIXTURE_<n>` marker so test recovery is verifiable.
- **Vitest unit tests** in `core/test/` plus headless smoke tests for the
  desktop/web bundles in `tests/`.
- **GitHub Actions release workflow** (`.github/workflows/release.yml`) that
  builds web + Linux + Windows + macOS + Android artifacts on tag push and
  attaches them to a GitHub Release.
- `CHANGELOG.md` and a platform-support matrix in `README.md`.

### Changed
- `web/index.html` is now the actual extractor SPA, not the README-viewer
  landing page (the previous pages workflow continues to publish from `web/`,
  so the live URL now serves the app).
- All Delphi sources moved from `src/` to `src/legacy-delphi/` and are no
  longer the primary build target.

### Notes
- iOS release builds aren't run in CI — they require an Apple Developer
  cert. See `tests/MANUAL.md`.
- macOS `.dmg` is unsigned. Users will need to right-click → Open the first
  time, or run `xattr -dr com.apple.quarantine /Applications/"Corrupt Office Extractor.app"`.

## [1.x]

Original Delphi 7 application (history preserved in
[`src/legacy-delphi/`](src/legacy-delphi/) and pre-2.0 commits).
