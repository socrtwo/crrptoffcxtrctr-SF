# Manual test plan: Android & iOS

CI builds an unsigned Android debug APK; iOS requires an Apple Developer
account, so its build is documented here and not automated. Both platforms
share the same `web/` bundle as engine, so the only thing being verified
manually is "does the WebView wrapper load and run the SPA correctly?"

## Prerequisites

- Node 20+, npm 10+
- For Android: JDK 21, Android Studio (or just the command-line tools), a
  device or emulator with API 26+
- For iOS: macOS, Xcode 15+, an iPhone simulator from Xcode's Devices window

Generate the corrupt fixtures and the web bundle first:

```sh
npm ci
npm run fixtures
npm run build:web
```

## Android (debug APK)

```sh
cd mobile
# First time only — creates the android/ native project:
npx cap add android
# Or, if android/ already exists:
npm run sync
cd android
./gradlew assembleDebug
```

The APK lands at `mobile/android/app/build/outputs/apk/debug/app-debug.apk`.

**Test plan (one device / emulator):**

1. Install the APK: `adb install app-debug.apk`
2. Launch *Corrupt Office Extractor*.
3. For each fixture in `tests/fixtures/`, copy it to the device storage
   (`adb push tests/fixtures/truncated.docx /sdcard/Download/`) and open it
   from the in-app file picker.
4. Confirm the recovered text panel contains `HELLO_FIXTURE_<n>` for each
   fixture (1 through 4).
5. Verify the entries panel lists `word/document.xml` /
   `xl/sharedStrings.xml` / `ppt/slides/slide1.xml` as appropriate.

## iOS (unsigned simulator build, no .ipa)

```sh
cd mobile
npx cap add ios          # first time only
npm run sync
cd ios/App
# Build for the simulator (no codesigning required):
xcodebuild \
  -workspace App.xcworkspace \
  -scheme App \
  -configuration Debug \
  -sdk iphonesimulator \
  -derivedDataPath build \
  CODE_SIGNING_ALLOWED=NO build
```

The resulting `.app` is at
`mobile/ios/App/build/Build/Products/Debug-iphonesimulator/App.app`.

**Test plan (one simulator):**

1. Install the app: `xcrun simctl install booted App.app`
2. Launch: `xcrun simctl launch booted com.socrtwo.crrptoffcxtrctr`
3. Drag each fixture from `tests/fixtures/` onto the simulator window to
   add it to the Files app, then open it from the in-app file picker.
4. Confirm `HELLO_FIXTURE_<n>` appears in the recovered-text panel for
   fixtures 1–4.

## Producing a signed `.ipa` (out of CI scope)

Requires an Apple Developer Program membership. From `mobile/ios/App`:

```sh
xcodebuild \
  -workspace App.xcworkspace \
  -scheme App \
  -configuration Release \
  -sdk iphoneos \
  -archivePath build/App.xcarchive archive

xcodebuild \
  -exportArchive \
  -archivePath build/App.xcarchive \
  -exportOptionsPlist exportOptions.plist \
  -exportPath build/ipa
```

Provide your `exportOptions.plist` with the Team ID and provisioning
profile of your choice; the resulting `.ipa` is at `build/ipa/App.ipa`.
