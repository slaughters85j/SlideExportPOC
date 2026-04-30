# SlideExportPOC

A small macOS SwiftUI app that demonstrates programmatic slide generation by driving Keynote.app via JXA (JavaScript for Automation) executed in-process through OSAKit, with export to either native Keynote (`.key`) or Microsoft PowerPoint (`.pptx`).

This is a **proof of concept**, not a library. It exists to validate an architecture, document the gotchas, and serve as a reference implementation for anyone trying to do the same thing.

## What it demonstrates

- A non-sandboxed, hardened-runtime macOS SwiftUI app sending Apple Events to Keynote.app, in-process, via `OSAKit.OSAScript` — no shell subprocess, no `osascript` binary.
- Building a JXA script dynamically in Swift, with strict string escaping, and surfacing per-step errors back to the host app.
- Rendering a SwiftUI Charts view to a PNG via `ImageRenderer` and embedding it as an image on a Keynote slide.
- Exporting to `.pptx` (via Keynote's `export` verb) or saving as `.key` (via Keynote's `save` verb — they are different APIs; see below).
- Validating "is the genuine Apple Keynote installed?" via `SecRequirementCreateWithString` + `SecStaticCodeCheckValidity` against the requirement `anchor apple generic and identifier "com.apple.Keynote"`.

## Why

This was built to answer a single question for a separate project: **can OSAKit + JXA replace a fragile, image-incapable, in-house slide-export library** for a real macOS app's project-export feature? Answer: yes. The PoC successfully produces both `.key` and `.pptx` files with embedded images.

## Stack

- Swift 5 (project default `MainActor` isolation, approachable concurrency on)
- SwiftUI + Swift Charts
- OSAKit (in-process JavaScript-for-Automation execution)
- Security framework (codesign-based validation of the target app)
- macOS 14+ deployment target
- No third-party dependencies

## Build & run

Open `SlideExportPOC.xcodeproj` in Xcode 15+ and run. You'll need:

1. Apple's Keynote installed from the Mac App Store (the app validates this and disables export with a clear diagnostic if it's missing).
2. The first export attempt will trigger a system consent dialog: *"SlideExportPOC wants to control Keynote.app"*. Click **Allow**.

If consent has already been silently denied (e.g. from a build before the entitlement was correct), the failure alert offers a **Reset Permission & Quit** button that runs `tccutil reset AppleEvents <bundle-id>` and terminates the app, so the next launch gets a fresh prompt.

## Project layout

| File | Purpose |
|---|---|
| `SlideExportPOCApp.swift` | `@main` entry point. |
| `ContentView.swift` | Split layout — sample content on the left, export controls on the right. Hosts the export sheet, the success/permission/failure alerts, and the in-app permission-reset flow. |
| `ExportSelectionView.swift` | Modal sheet — slide-item toggles, format picker (Keynote/PowerPoint), optional template picker, Keynote-installed diagnostic. |
| `SlideItem.swift` | Enum of slide kinds + `ExportFormat` enum with the right Keynote format strings. |
| `Constants.swift` | Hardcoded sample title, body paragraph, and chart data. Swap these to test other shapes. |
| `ChartView.swift` | Reusable Swift Charts bar chart, used both on-screen and offscreen by the renderer. |
| `ChartRenderer.swift` | `ImageRenderer` → 1600×800 PNG @ 2× into the temp directory. |
| `KeynoteExporter.swift` | The OSAKit/JXA engine. Install validation, JXA script builder with per-step error labels, OSAKit execution, error decoding. The reference implementation. |
| `SlideExportPOC.entitlements` | App Sandbox = NO + Apple Events automation entitlement. |

## Gotchas this PoC documents

These are the things that cost time and aren't well-documented in any single place. If you're trying to do something similar, save yourself the debugging session.

### 1. Hardened Runtime + Apple Events: you need both an entitlement AND an Info.plist key

With `ENABLE_HARDENED_RUNTIME = YES` (which is on by default for new Xcode projects and is a prerequisite for notarization), an app that sends Apple Events needs **all** of:

- `com.apple.security.automation.apple-events = true` in the entitlements file.
- `NSAppleEventsUsageDescription` in the Info.plist (set via `INFOPLIST_KEY_NSAppleEventsUsageDescription` build setting when `GENERATE_INFOPLIST_FILE = YES`).
- A code-signed binary with `CODE_SIGN_ENTITLEMENTS` pointing at the entitlements file.

App Sandbox status is independent of this — both sandboxed and non-sandboxed apps need these when hardened runtime is on.

**Without the entitlement**, Hardened Runtime rejects Apple Events *before TCC ever sees them*. The call fails with `errAEEventNotPermitted` (-1743), no consent prompt is shown, and the app **never appears** in System Settings → Privacy & Security → Automation. `tccutil reset` doesn't help — there is no TCC entry to reset. This failure mode looks like a permission problem but isn't.

**Without the Info.plist key**, TCC has no description string to show in the consent dialog and silently denies even when the entitlement is set.

### 2. Keynote's `export` verb does not handle native `.key` output

Keynote's `KeynoteExportFormat` enum members are: `Microsoft PowerPoint`, `PDF`, `HTML`, `slide images`, `QuickTime movie`. There is no `Keynote` member.

```js
// Conversion (.pptx, .pdf, .html, etc.)
kn.export(doc, { as: "Microsoft PowerPoint", to: Path("/path/out.pptx") });

// Native .key output — different verb!
kn.save(doc, { in: Path("/path/out.key") });
```

Calling `kn.export(doc, { as: "Keynote", to: ... })` produces `-1700 errAECoercionFail` ("Can't convert types") because AppleScript can't coerce the string `"Keynote"` to the `KeynoteExportFormat` enum. The error message is unhelpfully generic, which is why we wrap each JXA step in a labeled `step(name, fn)` helper that re-throws with the step name attached.

### 3. `anchor apple` is too strict for Mac App Store-distributed Apple apps

If you want to verify "is this app genuinely from Apple?" via the Security framework, use:

```
anchor apple generic and identifier "com.apple.Keynote"
```

NOT:

```
anchor apple and identifier "com.apple.Keynote"
```

The strict `anchor apple` form requires Apple's specific *Apple Software Signing* leaf certificate, which Apple uses for system-pre-installed binaries and command-line tools — but **not** for the Mac App Store version of Apple's own apps, which are signed with the *Apple Mac OS Application Signing* leaf cert. Both chains terminate at Apple Root CA, but the strict form rejects the Mac App Store path. `anchor apple generic` matches any chain anchored to Apple Root CA via Apple-controlled intermediates, so it covers both.

The bundle-ID equality predicate is what locks out impersonators — Apple controls the `com.apple.*` namespace at notarization and Mac App Store ingestion, so a non-Apple developer can't ship an app with `com.apple.Keynote` as its identifier through any normal channel.

Don't pin the Team Identifier — Apple uses several across iWork releases and across product lines. It's the wrong axis to lock against.

### 4. JXA error messages elide location information

OSAKit returns a single error dictionary for the whole script. If your JXA blows up at line 47 of a 60-line script, you get a message like *"Error: Error: Can't convert types."* with no line number. Wrap each major step in:

```js
function step(name, fn) {
    try { return fn(); }
    catch (e) {
        var msg = (e && e.message) ? e.message : String(e);
        throw new Error("[step:" + name + "] " + msg);
    }
}
```

…and you'll get `[step:writeOutput] Can't convert types.` instead, which actually points at the broken call.

### 5. Apple Events permission is per-codesign, not per-installation

If you change anything that affects the app's code signature (entitlements, Info.plist keys, signing identity) between builds, TCC may need to re-prompt. Conversely, if a previous build silently denied due to the issues above, the cached state can persist across builds. The in-app **Reset Permission & Quit** button runs `tccutil reset AppleEvents <bundle-id>` and terminates the app so the next launch starts fresh.

## Status

PoC validated end-to-end on macOS 14+. Not maintained as a library — fork it, copy from it, or use it as a reference for the patterns above.

## License

MIT — do whatever you want with it.
