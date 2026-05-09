<!--MODERNIZED:v2-->
# Crrptoffcxtrctr

> Originally a Windows-only Delphi 7 utility — now a cross-platform recovery tool that runs everywhere.

[![Live page](https://img.shields.io/badge/live-page-ff2e93?style=for-the-badge)](https://socrtwo.github.io/crrptoffcxtrctr-SF/)
[![Releases](https://img.shields.io/github/v/release/socrtwo/crrptoffcxtrctr-SF?style=for-the-badge&color=7c3aed)](https://github.com/socrtwo/crrptoffcxtrctr-SF/releases)
[![License](https://img.shields.io/github/license/socrtwo/crrptoffcxtrctr-SF?style=for-the-badge&color=22d3ee)](https://github.com/socrtwo/crrptoffcxtrctr-SF/blob/main/LICENSE)

🌐 **Live SPA (try it now):** https://socrtwo.github.io/crrptoffcxtrctr-SF/
📦 **Downloads:** [Releases](https://github.com/socrtwo/crrptoffcxtrctr-SF/releases)

---

Extracts text and embedded images from corrupt DOCX, XLSX, and PPTX files
that Office cannot open. The recovery engine walks ZIP local file headers
directly, so it works even when the central directory is missing or
truncated and even when individual entries fail their CRC.

The same TypeScript core powers every release: web SPA, Electron desktop
app, and Capacitor mobile wrapper — your file is never uploaded anywhere.

## Platform support matrix

| Platform | Format | Engine | How it's built |
| --- | --- | --- | --- |
| **Web (any browser)** | hosted SPA + installable PWA | TS core + esbuild bundle | `npm run build:web` |
| **Windows 10/11 (x64)** | `.exe` installer + portable `.zip` | Electron + TS core | `npm run build:win -w desktop` |
| **macOS 12+ (Intel & Apple Silicon)** | unsigned `.dmg` | Electron + TS core | `npm run build:mac -w desktop` |
| **Linux (Ubuntu / Debian / Fedora)** | `.AppImage` + `.deb` | Electron + TS core | `npm run build:linux -w desktop` |
| **ChromeOS** | install the PWA, or sideload the `.deb` under Crostini | Same Linux package + PWA | (see above) |
| **Android 8+ (API 26+)** | unsigned debug `.apk` | Capacitor + TS core | `npm run build:android-debug -w mobile` |
| **iOS 15+** | Xcode project (signed `.ipa` is your responsibility) | Capacitor + TS core | see [`tests/MANUAL.md`](tests/MANUAL.md) |

The legacy Delphi 7 source (Windows only) is preserved under
[`src/legacy-delphi/`](src/legacy-delphi/) for historical reference.

## Install / run per platform

### Use the web app

Open https://socrtwo.github.io/crrptoffcxtrctr-SF/ in any modern browser,
drag a corrupt file onto the page, and the recovered text + images appear
inline. Click *Install* in your browser to add it as a PWA.

### Windows / macOS / Linux desktop

Grab the matching artifact from the [Releases](https://github.com/socrtwo/crrptoffcxtrctr-SF/releases) page:

```
crrptoffcxtrctr-Setup-2.0.0.exe       # Windows installer
crrptoffcxtrctr-2.0.0-portable.zip    # Windows portable
Corrupt Office Extractor-2.0.0.dmg    # macOS (unsigned — see note)
crrptoffcxtrctr-2.0.0.AppImage        # Linux portable
crrptoffcxtrctr_2.0.0_amd64.deb       # Debian / Ubuntu / ChromeOS Crostini
```

> **macOS unsigned note:** First launch will be blocked by Gatekeeper.
> Right-click the app → *Open* to bypass, or run
> `xattr -dr com.apple.quarantine "/Applications/Corrupt Office Extractor.app"`.

### Android

Download `app-debug.apk` from the release, enable *Install unknown apps*
for your file manager, and tap to install. The build is unsigned (debug),
so it's intended for developer use; sign and ship a release variant from
the `mobile/android/` project for distribution.

### iOS

There is no prebuilt `.ipa` — building for iOS requires an Apple Developer
Program membership. Open the `mobile/ios/App/App.xcworkspace` Xcode project
on macOS and follow [`tests/MANUAL.md`](tests/MANUAL.md).

### ChromeOS

Two routes: install the web app as a PWA (recommended), or under Crostini
run `sudo apt install ./crrptoffcxtrctr_2.0.0_amd64.deb` after enabling
the Linux container.

## Build from source

```sh
git clone https://github.com/socrtwo/crrptoffcxtrctr-SF
cd crrptoffcxtrctr-SF
npm ci
npm run fixtures      # generate corrupt test files
npm test              # run the core test suite
npm run build:core    # compile the TS core
npm run build:web     # bundle the web SPA into dist/web/
npm run build -w desktop   # build Electron artifacts (current OS)
```

## Testing with corrupt samples

`scripts/make-fixtures.mjs` programmatically builds four deliberately-broken
files in `tests/fixtures/`, each carrying a recoverable
`HELLO_FIXTURE_<n>` marker:

| Fixture | Damage | Marker |
| --- | --- | --- |
| `truncated.docx` | last 2 KB chopped (kills central directory) | `HELLO_FIXTURE_1` |
| `bad-central-dir.xlsx` | EOCD + central directory zeroed | `HELLO_FIXTURE_2` |
| `crc-mismatch.pptx` | one byte flipped inside an entry's compressed data | `HELLO_FIXTURE_3` |
| `mixed-garbage.docx` | 1 KB of random bytes prepended to the file | `HELLO_FIXTURE_4` |

Verify all targets at once with:

```sh
npm run fixtures
npm run build:core
npm run build:web
node tests/verify-all.mjs
```

This runs the core, web, and Linux desktop engines against every fixture
and prints a pass/fail/manual matrix. The Playwright-driven browser test
at `tests/web-smoke.mjs` and the Android/iOS rows are documented in
[`tests/MANUAL.md`](tests/MANUAL.md).

## Repository layout

```
core/             TypeScript recovery engine (workspaces)
web-app/          esbuild script that bundles core into web/lib/
web/              Static SPA (index.html, manifest, service worker, icon)
desktop/          Electron main + preload, electron-builder config
mobile/           Capacitor wrapper for Android & iOS
scripts/          make-fixtures.mjs
src/legacy-delphi/  Original Delphi 7 sources, untouched
tests/            verify-all.mjs, web-smoke.mjs, MANUAL.md, fixtures/
.github/workflows/release.yml   Cross-platform release build
```

## Origin

This project began on **SourceForge** as a Windows-only Delphi 7 utility
([crrptoffcxtrctr](https://sourceforge.net/projects/crrptoffcxtrctr/)) and
was migrated to GitHub via [SF2GH Migrator](https://github.com/socrtwo/sf-to-github).
The 2.0.0 modernization rewrote the engine in TypeScript so the same code
runs on web, desktop, and mobile. The Delphi sources remain under
`src/legacy-delphi/` for historical reference.

## License

MIT — see [LICENSE](LICENSE).
