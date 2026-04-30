# SlideExportPOC

A small macOS SwiftUI app that demonstrates programmatic slide generation by driving Keynote.app via JXA (JavaScript for Automation) executed in-process through OSAKit, with export to either native Keynote (`.key`) or Microsoft PowerPoint (`.pptx`).

This is a **proof of concept**, not a library. It exists to validate an architecture, document the gotchas, and serve as a reference implementation for anyone trying to do the same thing.

## What it demonstrates

- A non-sandboxed, hardened-runtime macOS SwiftUI app sending Apple Events to Keynote.app, in-process, via `OSAKit.OSAScript` — no shell subprocess, no `osascript` binary.
- Building a JXA script dynamically in Swift, with strict string escaping, and surfacing per-step errors back to the host app.
- Rendering a SwiftUI Charts view to a PNG via `ImageRenderer` and embedding it as an image on a Keynote slide.
- Exporting to `.pptx` (via Keynote's `export` verb) or saving as `.key` (via Keynote's `save` verb — they are different APIs; see below).
- Output aspect ratio selection: **Wide (16:9, 1920×1080)** or **Standard (4:3, 1024×768)** — Keynote's own preset names.
- Custom template support — two routes:
  - **Installed Keynote themes** (recommended for `.kth`): the app queries `kn.themes` and lets the user pick a theme by name. Document is created via `kn.Document({documentTheme: kn.themes["…"], ...})`. No file open, no Keynote install dialog.
  - **Presentation file (`.key` / `.pptx`)**: opened directly with `kn.open(Path(...))`; your slides are appended to whatever's already there.
  - PowerPoint `.potx` is detected up-front and rejected with a remediation message — Keynote scripting can't consume it; the README documents both the manual workflow that works today and the OOXML post-processing path for a polished automated pipeline.
- A placeholder-discovery experiment for the chart slide that demonstrates one path toward template-aware image placement (see Item 4 in "Gotchas" below).
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

### 6. Custom templates: two routes, depending on the file type

The app surfaces two ways to pick a template via a "Choose…" Menu in the export sheet:

- **Installed Keynote Theme…** — lists every theme currently registered with Keynote (built-in + user-installed) by querying `kn.themes`. Pick one and the exporter creates a new document via `kn.Document({documentTheme: kn.themes["<name>"]})`. **No file open, no install dialog.** This is the right path for `.kth` themes.
- **Presentation File (.key / .pptx)…** — `NSOpenPanel` for an actual presentation file. The exporter calls `kn.open(Path(<file>))` and your slides are appended to whatever the file already contains.

`.kth` files cannot be picked from the file panel directly. `kn.open(Path("file.kth"))` triggers Keynote's modal "Add to Theme Chooser" install dialog **every time** — even after the user accepts it. Routing `.kth` through the installed-theme list bypasses that dialog entirely.

Workflow for a custom `.kth`:

1. Author your `.kth` in Keynote (File → Save as Theme…) or download one.
2. Double-click it in Finder. Keynote prompts "Add to Theme Chooser" — click it.
3. In SlideExportPOC, open the export sheet and pick **Choose… → Installed Keynote Theme…**. Your theme will be in the list. Pick it.
4. Run the export. The document is created against your theme; the master-slide names from your theme drive layout selection.

If you accidentally pass a `.kth` path through the file picker (say, programmatically), the app surfaces a clear error pointing at the installed-theme path.

Things to know about templates generally:

- When a template is supplied, the **template's slide dimensions win** — the aspect picker is informational only and the export sheet greys it out with a note.
- `.installedTheme` paths produce a single auto-created starter slide; the script deletes that after appending its own content (matching the no-template flow).
- `.file(.key)` and `.file(.pptx)` paths carry the user's existing slides; the script **does not** delete them — your slides are appended after the originals.
- For custom themes whose master-slide names don't match Apple's standard conventions ("Title, Content", "Photo - Horizontal", etc.), the master-slide picker falls back to `masters[0]`. The script also writes the resolved master names into a JSON diagnostics object that's pretty-printed to the Xcode console after every successful run, so the available masters are visible without breakpointing.

### 7. PowerPoint `.potx` templates: three paths, one of them implemented

Keynote's scripting interface has no API to consume PowerPoint `.potx` files. So `.potx` support requires bypassing Keynote — and this PoC implements one of the two practical paths so DA can reuse the pattern.

**Path A — Manual workflow (works without writing any code).** Export `.pptx` vanilla from SlideExportPOC, open in PowerPoint, **Design → Themes → Browse for Themes…**, pick your `.potx`. PowerPoint applies the template's masters, colors, fonts, and layouts to your existing slides. Confirmed clean. Best for one-offs.

**Path B — `.potxOverlay` (implemented in `PPTXTemplateMerger.swift`).** A pure-Swift OOXML pipeline: copy the `.potx` as the base, flip the package MIME from `presentationml.template.main+xml` → `presentationml.presentation.main+xml`, inject content into the template's existing slide placeholders, re-zip as `.pptx`. **Doesn't require Keynote at all** — no Apple Events, no install dialogs, no app launch.

In the export sheet, choose **Choose… → PowerPoint Template (.potx) overlay…** (only visible when Format = PowerPoint). The merger:

- Resolves slide order by numeric suffix on `slide*.xml`.
- Finds `<p:ph type="title"/>` on slide 1 and replaces its text body with the title string.
- Finds `<p:ph type="body" idx="1"/>` on slide 2, copies the chart PNG into `ppt/media/`, registers a relationship in `ppt/slides/_rels/slide2.xml.rels`, replaces the body shape with a `<p:pic>` element sized to the body's geometry.
- Geometry resolution uses a three-tier fallback: layout body `<a:xfrm>` → slide-relative frame from `<p:sldSz>` (8% horizontal / 13% vertical margins) → hardcoded defaults. The middle tier handles the common case where layouts inherit body geometry from the master rather than declaring it directly (true on the sample template).

Implementation is ~500 LOC of `Foundation.XMLDocument` plus `/usr/bin/zip` and `/usr/bin/unzip`. No third-party deps.

Findings worth knowing for a production implementation in DA:

- The `.potx` you author defines the binding contract. The PoC merger assumes "slide 1 has a title placeholder → put title there; slide 2 has a body placeholder → put image there." A more flexible mapping in production would let the template name slides explicitly (e.g. `<p:cNvPr name="Title Slide"/>` / `<p:cNvPr name="Chart Slide"/>`) and have the binding code look up by name rather than slide order.
- Empty placeholders show layout-level prompt text (e.g. "TITLE", "Click to add subtitle") in PowerPoint's normal/edit view but not in slideshow. If you see prompt text on a slide where you didn't inject content, that's why — it's not actual content.
- The `<p:pic>` we emit replaces the body `<p:sp>` outright, which removes the placeholder semantics on that slide. PowerPoint still renders the layout's body placeholder beneath as a design hint outline in edit view; in slideshow the picture covers it. If keeping the placeholder semantics matters, the merger would need to insert the `<p:pic>` *alongside* the body shape and either leave the body empty or move the body off-canvas.
- Slide dimensions in the user's sample `.potx` are non-standard (`cx=24384000 cy=13716000`, ≈26.7"×15") — a Keynote-export style sizing. The slide-relative geometry fallback handles arbitrary canvas sizes correctly.

**Path C — Theme-only swap (not implemented, easy if needed).** Run a vanilla `.pptx` export through Keynote, then post-process: replace just `ppt/theme/theme1.xml`, `ppt/slideMasters/`, and `ppt/slideLayouts/` from the `.potx` into the export, rewrite slide-to-layout relationships, re-zip. Lower ceiling than Path B (Keynote's positioned shapes don't move to match new layouts) but useful if the goal is "Keynote-driven structure, .potx-driven styling."

For DA: Path B is the one to look at. Authoring one canonical `.potx` (and one canonical `.kth` for Keynote) with deterministic placeholder structure converts the export problem into a deterministic content-injection transform. The PoC's `PPTXTemplateMerger.swift` is the reference implementation.

### 7. Cross-template image placement: text is template-agnostic, images are not (without effort)

This is the question motivating the "polish, minimal-edit" goal. Findings from the `chartSlide` experiment in `KeynoteExporter.swift`:

- **Text** binds cleanly to the master's `defaultTitleItem` and `defaultBodyItem`. On `.pptx` export Keynote maps these to PowerPoint's Title and Body placeholders, so typography, color, and position are controlled by the chosen template. Changing templates "just works" for title + body content.
- **Images** are harder. Keynote's scripting model exposes `slide.images` (free-floating images you've added) and `slide.iWorkItems()` (everything on the slide), but does **not** expose a clean "drop my PNG into the master's image placeholder" API analogous to `defaultTitleItem`. PowerPoint's OOXML *does* have named picture placeholders, but Keynote's `.pptx` exporter generally rasterizes positioned images to absolute coordinates rather than emitting `<p:ph type="pic">` nodes. Round-tripping placeholder semantics through Keynote → `.pptx` is unreliable.

The PoC implements a best-effort discovery pass: on the chart slide, after instantiating from a master (preferring "Photo - Horizontal" / "Photo" / "Blank"), it iterates `slide.iWorkItems()` looking for an image-class item large enough to be the master's hero placeholder. If found, the chart PNG steals that item's position and size and the original is deleted; if not, the chart falls back to an aspect-relative frame computed in Swift (≈8% horizontal, ≈13% vertical margins). Whether the placeholder pass actually finds anything depends entirely on the chosen master and how the template was authored.

**Recommended pattern for production code** (e.g. a polished slide export in a real app): author **one canonical `.kth`** with deterministic master-slide names and known placeholder positions. In code, hardcode the mapping from app data to those known slots. This converts "discover what's there at runtime" into "place exactly here, every time." Multiple official templates (light, dark, corporate) can share the same master-slide names so the binding code stays the same — only the visual styling changes. This is the path the SlideExportPOC documents but does **not** itself implement, since one canonical template was out of scope.

## Status

PoC validated end-to-end on macOS 14+. Not maintained as a library — fork it, copy from it, or use it as a reference for the patterns above.

## License

MIT — do whatever you want with it.
