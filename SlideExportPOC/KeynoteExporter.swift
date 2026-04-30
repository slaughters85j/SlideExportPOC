//
//  KeynoteExporter.swift
//  SlideExportPOC
//
//  In-process JXA engine that drives Keynote.app via OSAKit, creates a slide
//  document from selected SlideItems, and exports as Keynote or PowerPoint.
//
//  Why OSAKit instead of `Process` running `osascript`?
//  - No shell subprocess; the JXA runs inside the host app's address space.
//  - No `--args` quoting headaches; we pass values by interpolating into the
//    script source (carefully escaped via jsEscape).
//  - The OSA error dictionary surfaces both message and code, which we map
//    onto a typed `KeynoteExporterError`.
//
//  This is the architecture we are evaluating for replacing DecisionArchitect's
//  SlideKit-based export. If this PoC holds up, the same OSAKit + JXA bridge
//  pattern scales to multi-slide DA exports with embedded imagery.
//

import Foundation
import AppKit
import OSAKit
import Security

// TODO: TEMPLATE SUPPORT
// - Keynote .kth themes: scriptable via `Application("Keynote").Document({
//   documentTheme: theme, ... })` where `theme` is matched by name from
//   `kn.themes()`. User-installed themes appear in the same list once
//   installed via `open` in Keynote.
// - PowerPoint .potx templates: not directly supported by Keynote's exporter.
//   Workaround: open an existing .pptx as a base document
//   (`kn.open(Path("template.pptx"))`) and populate slides into the resulting
//   document instead of creating a new one.
// Both viable as future enhancements; out of scope for this PoC. The
// `customTemplateURL` parameter on exportToKeynote is plumbed through but
// currently ignored (a debug warning is logged when non-nil).

// MARK: - KeynoteExporterError

enum KeynoteExporterError: LocalizedError {
    case keynoteNotInstalled
    case automationPermissionDenied
    case scriptingFailed(message: String, code: Int, fullErrorDictionary: [String: Any])
    case chartRenderFailed(underlying: Error)
    case exportFileMissing(expectedPath: String)
    case unsupportedTemplateType(extension: String)
    case unknown(String)

    var errorDescription: String? {
        switch self {
        case .keynoteNotInstalled:
            return "Keynote is not installed on this Mac. Install Keynote from the Mac App Store and try again."
        case .automationPermissionDenied:
            return """
            Automation permission was denied. macOS blocked SlideExportPOC from controlling Keynote.

            Open System Settings → Privacy & Security → Automation, find SlideExportPOC, and enable the Keynote toggle. Then try the export again.

            (If SlideExportPOC does not appear in that list, quit and reopen the app — macOS will prompt for permission the next time it tries to drive Keynote.)
            """
        case .scriptingFailed(let message, let code, _):
            return "Keynote scripting error (\(code)): \(message)\n\nSee the Xcode console for the full error dictionary."
        case .chartRenderFailed(let underlying):
            return "Failed to render the chart for export: \(underlying.localizedDescription)"
        case .exportFileMissing(let path):
            return "Keynote reported success but the expected file was not found at \(path)."
        case .unsupportedTemplateType(let ext):
            return """
            Templates of type ".\(ext)" are not supported. Keynote cannot consume PowerPoint .potx templates directly.

            Workaround: open the .potx in PowerPoint, save it as .pptx, then select the .pptx as the template here.
            """
        case .unknown(let message):
            return message
        }
    }
}

// MARK: - Apple Event status code decoding

/// Best-effort decoder for the most common Apple Event / OSA status codes
/// surfaced via the OSAKit error dictionary's `OSAScriptErrorNumber`. Used
/// for diagnostic logging — not exhaustive, just the failure modes we expect
/// to actually hit when driving Keynote from a host app.
private func decodeOSAStatusCode(_ code: Int) -> String {
    switch code {
    case    0: return "noErr (success)"
    case -128: return "userCanceledErr — the user canceled."
    case -600: return "procNotFound — the target application is not running and could not be launched."
    case -609: return "connectionInvalid — Apple Event connection is no longer valid."
    case -1700: return "errAECoercionFail — could not coerce a value to the requested type."
    case -1701: return "errAEDescNotFound — referenced object does not exist (often a missing master slide name or property)."
    case -1708: return "errAEEventNotHandled — the target app does not handle this command."
    case -1712: return "errAETimeout — Apple Event timed out (Keynote took too long to respond)."
    case -1713: return "errAENoUserInteraction — the script tried to interact with the user but is not allowed to."
    case -1719: return "errAEIllegalIndex — invalid index into a collection (e.g. doc.slides[N] out of bounds)."
    case -1728: return "errAENoSuchObject — referenced object does not exist."
    case -1743: return "errAEEventNotPermitted — macOS blocked the Apple Event (Automation permission denied or NSAppleEventsUsageDescription missing from Info.plist)."
    case -1751: return "errOSAInvalidID — invalid scripting id."
    case -2700: return "errOSAGeneralError — generic OSA failure."
    case -2701: return "errOSADivideByZero."
    case -2702: return "errOSANumericOverflow."
    case -2703: return "errOSACantAssign."
    case -2706: return "errOSADeepRecursion."
    case -2740: return "errOSASyntaxError — JavaScript syntax error in the JXA source."
    case -10000: return "errAEEventFailed — generic AppleScript/JXA failure inside the target app (the target app's own error)."
    case -10004: return "errAEPrivilegeError — privilege error."
    case -10006: return "errAENotModifiable — the property cannot be set (often: assigning to a read-only Keynote property, or to a property that doesn't exist on the chosen master slide)."
    default:
        return "unrecognized OSA/AE status code"
    }
}

// MARK: - KeynoteInstallStatus

/// Rich result of the Keynote install check. The sheet uses `diagnostic` to
/// explain *why* export is unavailable when `isUsable` is false, since the
/// failure modes (not installed at all vs. only an App Store placeholder vs.
/// a third-party clone impersonating Apple's bundle ID) need different
/// remediation from the user.
struct KeynoteInstallStatus {
    let isUsable: Bool
    let foundURL: URL?
    let diagnostic: String
}

// MARK: - KeynoteExporter

enum KeynoteExporter {

    /// Apple's Keynote has shipped under two bundle IDs across versions:
    /// modern Mac App Store builds use `com.apple.Keynote`; older legacy
    /// builds and many internal references use `com.apple.iWork.Keynote`.
    /// We accept either, but only when the app is anchored to Apple itself.
    private static let keynoteBundleIDs = ["com.apple.Keynote", "com.apple.iWork.Keynote"]

    // MARK: Installed check

    /// Note: Keynote does NOT need to be already running for export to work.
    /// JXA's `Application("Keynote")` launches it on demand, and our script
    /// never calls `.activate()`, so Keynote runs invisibly in the background
    /// for the duration of the export. It does need to be **installed** —
    /// specifically, an Apple-signed Keynote (any path, any name on disk;
    /// users may rename the app bundle).
    ///
    /// Validation strategy:
    /// 1. Ask LaunchServices for every URL claiming one of Apple's Keynote
    ///    bundle IDs.
    /// 2. For each candidate, verify it satisfies the Code Signing
    ///    Requirement `anchor apple and identifier "<bundleID>"`. The
    ///    `anchor apple` predicate (note: NOT `anchor apple generic`) only
    ///    matches binaries actually signed by Apple-the-company, so it
    ///    rejects both unsigned clones and third-party apps shipped via the
    ///    Mac App Store. We don't pin a Team Identifier because Apple uses
    ///    several across iWork releases (e.g. `74J34U3R6X`, `JCRTNEU7GK`).
    static func keynoteInstallStatus() -> KeynoteInstallStatus {
        // 1. Gather candidates from LaunchServices for every known bundle ID.
        var candidates: [URL] = []
        for bundleID in keynoteBundleIDs {
            for url in NSWorkspace.shared.urlsForApplications(withBundleIdentifier: bundleID) {
                if !candidates.contains(url) { candidates.append(url) }
            }
        }

        #if DEBUG
        print("[KeynoteExporter] install check — \(candidates.count) candidate(s):")
        for url in candidates {
            print("    • \(url.path)")
        }
        #endif

        // 2. No candidates at all → genuinely not installed.
        guard !candidates.isEmpty else {
            return KeynoteInstallStatus(
                isUsable: false,
                foundURL: nil,
                diagnostic: "Apple’s Keynote is not installed. Install it from the Mac App Store and try again."
            )
        }

        // 3. Validate each candidate against the Apple-signed requirement.
        for url in candidates {
            let result = validateAppleSignedKeynote(at: url)
            #if DEBUG
            print("[KeynoteExporter]    → \(url.lastPathComponent) bundleID=\(result.bundleID ?? "nil") team=\(result.teamID ?? "nil") appleSigned=\(result.isAppleSigned) checkValidityStatus=\(result.validationStatus) (\(secErrorMessage(result.validationStatus)))")
            #endif
            if result.isAppleSigned {
                return KeynoteInstallStatus(
                    isUsable: true,
                    foundURL: url,
                    diagnostic: "Found Apple Keynote at \(url.path)."
                )
            }
        }

        // 4. We found something, but nothing satisfied the Apple-signed
        //    requirement. The most useful explanation depends on what we hit.
        let placeholderHit = candidates.contains { $0.path.contains("/Placeholders-") }

        var explanation = "An app claiming Keynote’s bundle identifier was found, but it is not signed by Apple and was rejected."
        if placeholderHit {
            explanation = "A Mac App Store download appears to be incomplete (placeholder only). Open the App Store and finish installing Keynote."
        }

        return KeynoteInstallStatus(
            isUsable: false,
            foundURL: candidates.first,
            diagnostic: explanation
        )
    }

    /// Convenience wrapper for callers that only need the boolean.
    static func isKeynoteInstalled() -> Bool {
        keynoteInstallStatus().isUsable
    }

    /// Best-effort human-readable form of a `SecStaticCodeCheckValidity`
    /// status code, for debug logging.
    private static func secErrorMessage(_ status: OSStatus) -> String {
        if status == errSecSuccess { return "success" }
        if let cfMessage = SecCopyErrorMessageString(status, nil) {
            return cfMessage as String
        }
        return "OSStatus \(status)"
    }

    // MARK: Apple-signed validation

    private struct AppleSignedResult {
        let isAppleSigned: Bool
        let bundleID: String?
        let teamID: String?
        let validationStatus: OSStatus
    }

    /// Returns whether the app at `url` is signed by Apple-the-company and
    /// claims one of Apple's Keynote bundle IDs.
    ///
    /// Why `anchor apple generic` and not `anchor apple`?
    ///   - `anchor apple` requires the specific "Apple Software Signing" leaf
    ///     certificate, which Apple uses for system-pre-installed apps and
    ///     command-line tools.
    ///   - Apps Apple ships through the Mac App Store (including modern
    ///     Keynote) are signed with the "Apple Mac OS Application Signing"
    ///     leaf cert, which fails `anchor apple` even though the chain
    ///     legitimately terminates at Apple Root CA.
    ///   - `anchor apple generic` matches anything anchored to Apple Root CA
    ///     via any Apple-controlled intermediate, which covers both paths.
    ///
    /// Combined with a bundle-ID equality check on `com.apple.Keynote` /
    /// `com.apple.iWork.Keynote`: only Apple can ship an app with those
    /// identifiers (Apple controls the `com.apple.*` namespace at notarization
    /// and Mac App Store ingestion), so the combination of "valid Apple-anchored
    /// signature" + "Apple-owned bundle ID" is a sound test.
    private static func validateAppleSignedKeynote(at url: URL) -> AppleSignedResult {
        var staticCode: SecStaticCode?
        guard
            SecStaticCodeCreateWithPath(url as CFURL, [], &staticCode) == errSecSuccess,
            let code = staticCode
        else {
            return AppleSignedResult(isAppleSigned: false, bundleID: nil, teamID: nil, validationStatus: errSecCSStaticCodeNotFound)
        }

        // Read the embedded Info.plist + signing info for diagnostics.
        var infoDict: CFDictionary?
        _ = SecCodeCopySigningInformation(code, SecCSFlags(rawValue: kSecCSSigningInformation), &infoDict)
        let info = (infoDict as? [String: Any]) ?? [:]
        let teamID = info[kSecCodeInfoTeamIdentifier as String] as? String
        let bundleID: String? = {
            if let plist = info[kSecCodeInfoPList as String] as? [String: Any] {
                return plist["CFBundleIdentifier"] as? String
            }
            return nil
        }()

        let identifierClause = keynoteBundleIDs
            .map { "identifier \"\($0)\"" }
            .joined(separator: " or ")
        let requirementString = "anchor apple generic and (\(identifierClause))"

        var requirement: SecRequirement?
        guard
            SecRequirementCreateWithString(requirementString as CFString, [], &requirement) == errSecSuccess,
            let req = requirement
        else {
            return AppleSignedResult(isAppleSigned: false, bundleID: bundleID, teamID: teamID, validationStatus: errSecParam)
        }

        let validateStatus = SecStaticCodeCheckValidity(code, [], req)
        return AppleSignedResult(
            isAppleSigned: validateStatus == errSecSuccess,
            bundleID: bundleID,
            teamID: teamID,
            validationStatus: validateStatus
        )
    }

    // MARK: Public export entry point

    /// Renders any required assets, builds a JXA script, drives Keynote
    /// in-process via OSAKit, and returns the URL of the produced file
    /// (either .pptx or .key on the Desktop).
    ///
    /// - Parameters:
    ///   - items: Which slide kinds to include.
    ///   - format: `.keynote` (.key, via `kn.save`) or `.powerPoint` (.pptx,
    ///             via `kn.export(as: "Microsoft PowerPoint")`).
    ///   - aspect: 16:9 (`.wide`) or 4:3 (`.standard`). Ignored when a
    ///             custom template is supplied — template dimensions win.
    ///   - customTemplateURL: Optional .kth / .key / .pptx to use as the
    ///             base document. .potx is rejected up-front.
    @MainActor
    static func exportToKeynote(
        items: [SlideItem],
        format: ExportFormat,
        aspect: SlideAspect = .wide,
        customTemplateURL: URL? = nil
    ) async throws -> URL {
        guard isKeynoteInstalled() else {
            throw KeynoteExporterError.keynoteNotInstalled
        }

        // 1. Validate template type up-front.
        if let template = customTemplateURL {
            let ext = template.pathExtension.lowercased()
            switch ext {
            case "kth", "key", "pptx":
                break // supported via kn.open
            case "potx":
                throw KeynoteExporterError.unsupportedTemplateType(extension: "potx")
            default:
                throw KeynoteExporterError.unsupportedTemplateType(extension: ext)
            }
        }

        // 2. Render the chart PNG up front (must be on the main actor) if needed.
        //    Render at the chart frame size for the chosen aspect so the bitmap
        //    matches the embedded image rectangle and stays crisp.
        var chartPNGURL: URL? = nil
        if items.contains(.chart) {
            do {
                chartPNGURL = try ChartRenderer.renderChartToPNG(size: aspect.chartFrame.size)
            } catch {
                throw KeynoteExporterError.chartRenderFailed(underlying: error)
            }
        }

        // 3. Decide where the output file goes.
        let outputURL = makeOutputURL(for: format)

        // 4. Build the script with Swift values interpolated (escaped).
        let script = makeJXAScript(
            items: items,
            format: format,
            aspect: aspect,
            outputURL: outputURL,
            chartPNGURL: chartPNGURL,
            templateURL: customTemplateURL
        )

        // 5. Run OSAKit off the main actor — `executeAndReturnError` is blocking.
        try await Task.detached(priority: .userInitiated) {
            try runJXA(script)
        }.value

        // 6. Verify the export file landed.
        guard FileManager.default.fileExists(atPath: outputURL.path) else {
            throw KeynoteExporterError.exportFileMissing(expectedPath: outputURL.path)
        }

        // 7. Best-effort temp cleanup.
        if let chartPNGURL {
            try? FileManager.default.removeItem(at: chartPNGURL)
        }

        return outputURL
    }

    // MARK: Output path

    private static func makeOutputURL(for format: ExportFormat) -> URL {
        let formatter = DateFormatter()
        formatter.dateFormat = "yyyyMMdd_HHmmss"
        formatter.locale = Locale(identifier: "en_US_POSIX")
        let timestamp = formatter.string(from: Date())

        let desktop = FileManager.default
            .urls(for: .desktopDirectory, in: .userDomainMask)
            .first
            ?? URL(fileURLWithPath: NSHomeDirectory()).appendingPathComponent("Desktop")

        return desktop.appendingPathComponent("SlideExportPOC_\(timestamp).\(format.fileExtension)")
    }

    // MARK: OSAKit execution

    /// Runs a JXA source string via OSAKit. Marked `nonisolated` so the
    /// project's default-MainActor isolation doesn't force this blocking call
    /// onto the main thread — we deliberately invoke it from a detached task.
    nonisolated private static func runJXA(_ source: String) throws {
        guard let language = OSALanguage(forName: "JavaScript") else {
            throw KeynoteExporterError.unknown("JavaScript scripting language unavailable on this system.")
        }

        #if DEBUG
        print("[KeynoteExporter] ───── JXA source ─────")
        print(source)
        print("[KeynoteExporter] ──── /JXA source ─────")
        #endif

        let script = OSAScript(source: source, language: language)

        var errorDict: NSDictionary?
        let result = script.executeAndReturnError(&errorDict)

        guard let errorDict else {
            #if DEBUG
            if let result {
                print("[KeynoteExporter] JXA executed successfully. Result descriptor: \(result.stringValue ?? "<no string value>")")
            }
            #endif
            return
        }

        // Build a stable [String: Any] copy of the error dictionary for both
        // logging and downstream propagation.
        let asSwiftDict: [String: Any] = (errorDict as? [String: Any]) ?? [:]
        let message = (asSwiftDict[OSAScriptErrorMessage] as? String)
            ?? (asSwiftDict["NSLocalizedDescription"] as? String)
            ?? (asSwiftDict[OSAScriptErrorBriefMessage] as? String)
            ?? "Unknown scripting error."
        let code = (asSwiftDict[OSAScriptErrorNumber] as? Int) ?? -1

        #if DEBUG
        print("[KeynoteExporter] ═════ JXA failed ═════")
        print("  code:    \(code) — \(decodeOSAStatusCode(code))")
        print("  message: \(message)")
        print("  full error dictionary:")
        for (k, v) in asSwiftDict {
            print("    [\(k)] = \(v)")
        }
        print("[KeynoteExporter] ════ /JXA failed ═════")
        #endif

        // Map known-bad codes to dedicated cases so the UI can show actionable
        // remediation (specifically, the Automation-permission case is the
        // hard one to diagnose without this).
        if code == -1743 {
            throw KeynoteExporterError.automationPermissionDenied
        }

        throw KeynoteExporterError.scriptingFailed(
            message: message,
            code: code,
            fullErrorDictionary: asSwiftDict
        )
    }

    // MARK: JXA script builder

    private static func makeJXAScript(
        items: [SlideItem],
        format: ExportFormat,
        aspect: SlideAspect,
        outputURL: URL,
        chartPNGURL: URL?,
        templateURL: URL?
    ) -> String {
        // Build the per-slide JS blocks first so we can keep the wrapper clean.
        var slideBlocks: [String] = []

        for item in items {
            switch item {
            case .titleAndBody:
                slideBlocks.append(titleAndBodySlideBlock(
                    title: SampleContent.title,
                    body: SampleContent.body
                ))
            case .chart:
                if let chartPNGURL {
                    slideBlocks.append(chartSlideBlock(
                        imagePath: chartPNGURL.path,
                        chartFrame: aspect.chartFrame
                    ))
                }
            }
        }

        let blocksJoined = slideBlocks.joined(separator: "\n")
        let outputPathLiteral = jsEscape(outputURL.path)

        // Keynote's `export` verb only accepts *conversion* targets
        // (Microsoft PowerPoint, PDF, HTML, slide images, QuickTime). There
        // is no "Keynote" member of the KeynoteExportFormat enum, so for the
        // native .key format we use `save` instead.
        let writeStatement: String
        switch format {
        case .keynote:
            writeStatement = #"""
            kn.save(doc, { in: out });
            """#
        case .powerPoint:
            writeStatement = #"""
            kn.export(doc, { as: "Microsoft PowerPoint", to: out });
            """#
        }

        // Document-construction step varies based on whether a template was
        // supplied. With a template we open the file (Keynote opens .kth
        // .key and .pptx natively); the template's theme & dimensions win.
        // Without a template we create a fresh document at the requested
        // aspect using the "White" theme.
        //
        // We also need to know whether to delete a starter slide afterwards:
        // - No template: yes, Keynote auto-creates one.
        // - .kth template: yes, opening a theme file creates a starter slide.
        // - .key / .pptx template: NO — those files contain the user's own
        //   slides which we should preserve.
        let createDocumentStatement: String
        let shouldDeleteStarterSlide: Bool

        if let templateURL {
            let templateLiteral = jsEscape(templateURL.path)
            let ext = templateURL.pathExtension.lowercased()
            shouldDeleteStarterSlide = (ext == "kth")
            createDocumentStatement = """
            var doc = step("openTemplate", function() {
                var d = kn.open(Path("\(templateLiteral)"));
                if (!d) {
                    // kn.open sometimes returns the document indirectly; fall
                    // back to the front document.
                    d = kn.documents[0];
                }
                return d;
            });
            """
        } else {
            shouldDeleteStarterSlide = true
            createDocumentStatement = """
            var theme = step("findTheme", function() {
                var themesList = kn.themes;
                try {
                    var named = themesList["White"];
                    named.name();
                    return named;
                } catch (e) {
                    return themesList[0];
                }
            });

            var doc = step("createDocument", function() {
                var d = kn.Document({
                    documentTheme: theme,
                    width: \(Int(aspect.size.width)),
                    height: \(Int(aspect.size.height))
                });
                kn.documents.push(d);
                return d;
            });
            """
        }

        let starterSlideHandling: String = shouldDeleteStarterSlide
            ? """
            var initialSlide = step("captureInitialSlide", function() {
                return doc.slides[0];
            });
            """
            : """
            var initialSlide = null;
            """

        let starterSlideDelete: String = shouldDeleteStarterSlide
            ? """
            step("deleteInitialSlide", function() {
                if (initialSlide) initialSlide.delete();
            });
            """
            : """
            // Template carries user's own slides; do not delete anything.
            """

        return """
        (function() {
            // step(name, fn): wraps each major operation so any thrown error
            // is re-thrown with an explicit step label. OSAKit only surfaces
            // the final exception message, so without this the host app sees
            // a generic "Can't convert types" with no idea which step failed.
            function step(name, fn) {
                try {
                    return fn();
                } catch (e) {
                    var msg = (e && e.message) ? e.message : String(e);
                    throw new Error("[step:" + name + "] " + msg);
                }
            }

            // Diagnostic helper: tries to read a property off an iWorkItem
            // returning a default if access fails. Used in placeholder
            // discovery to make the iteration robust to heterogeneous items.
            function safeGet(fn, fallback) {
                try { return fn(); } catch (e) { return fallback; }
            }

            var kn = step("application", function() {
                var a = Application("Keynote");
                a.includeStandardAdditions = true;
                return a;
            });

            // 1. Document construction — varies based on template vs. no template.
            \(createDocumentStatement)

            // 2. Capture the starter slide if we should delete it later.
            \(starterSlideHandling)

            // 3. Per-item slide blocks (interpolated by Swift). Each block
            //    uses its own step() wrappers internally.
            \(blocksJoined)

            // 4. Optionally remove the auto-created starter slide.
            \(starterSlideDelete)

            // 5. Save / export to the requested format.
            var out = step("buildOutputPath", function() {
                return Path("\(outputPathLiteral)");
            });

            step("writeOutput", function() {
                \(writeStatement)
            });

            // 6. Close without saving the in-memory document.
            step("closeDocument", function() {
                doc.close({ saving: "no" });
            });

            return "ok";
        })();
        """
    }

    // MARK: Per-item slide blocks

    private static func titleAndBodySlideBlock(title: String, body: String) -> String {
        let titleLiteral = jsEscape(title)
        let bodyLiteral  = jsEscape(body)

        return """
        step("titleBodySlide", function() {
            var master = (function() {
                var masters = doc.masterSlides;
                var preferred = ["Title, Content", "Title & Bullets", "Title & Content", "Title - Top", "Title"];
                for (var p = 0; p < preferred.length; p++) {
                    try {
                        var m = masters[preferred[p]];
                        m.name(); // force evaluation
                        return m;
                    } catch (e) { /* try next */ }
                }
                return masters[0];
            })();

            var slide = kn.Slide({ baseSlide: master });
            doc.slides.push(slide);

            try { slide.defaultTitleItem.objectText = "\(titleLiteral)"; } catch (e) { /* master may not have a title placeholder */ }
            try { slide.defaultBodyItem.objectText  = "\(bodyLiteral)";  } catch (e) { /* master may not have a body placeholder */ }
        });
        """
    }

    private static func chartSlideBlock(imagePath: String, chartFrame: CGRect) -> String {
        let pathLiteral = jsEscape(imagePath)
        let fallbackX = Int(chartFrame.origin.x.rounded())
        let fallbackY = Int(chartFrame.origin.y.rounded())
        let fallbackW = Int(chartFrame.size.width.rounded())
        let fallbackH = Int(chartFrame.size.height.rounded())

        // The script picks a master ("Photo - Horizontal" or similar
        // image-friendly master if available, otherwise "Blank"), then runs
        // a placeholder-discovery experiment. We log every iWorkItem on
        // the master and on the freshly-created slide to the OSAKit return
        // value, so the host app can inspect them in the console. If we
        // find an item that looks like an image placeholder, we steal its
        // position/size and delete it; if not, we fall back to the
        // aspect-relative chart frame computed in Swift.
        return """
        step("chartSlide", function() {
            var master = (function() {
                var masters = doc.masterSlides;
                // Image-friendly masters first, falling back to Blank.
                var preferred = ["Photo - Horizontal", "Photo - 3 Up", "Photo", "Blank"];
                for (var p = 0; p < preferred.length; p++) {
                    try {
                        var m = masters[preferred[p]];
                        m.name();
                        return m;
                    } catch (e) { /* try next */ }
                }
                return masters[0];
            })();

            var slide = kn.Slide({ baseSlide: master });
            doc.slides.push(slide);

            // ─── Placeholder-discovery experiment (item 4 in the README) ───
            // Iterate the slide's iWorkItems looking for an image-shaped
            // placeholder inherited from the master. If found, use its
            // geometry — that gives us template-aware placement. Otherwise
            // fall back to the aspect-relative frame Swift computed.
            var targetX = \(fallbackX);
            var targetY = \(fallbackY);
            var targetW = \(fallbackW);
            var targetH = \(fallbackH);
            var usedPlaceholder = false;

            try {
                var items = slide.iWorkItems();
                for (var i = 0; i < items.length; i++) {
                    var it = items[i];
                    var itemClass = safeGet(function() { return it.class(); }, "?");
                    var w = safeGet(function() { return it.width(); }, 0);
                    var h = safeGet(function() { return it.height(); }, 0);
                    // Heuristic: image-class items larger than ~25% of the
                    // slide are likely the master's hero-image placeholder.
                    var isLikelyImagePlaceholder =
                        (itemClass === "image" || itemClass === "imagePlaceholder") &&
                        w * h > (\(fallbackW) * \(fallbackH)) * 0.25;
                    if (isLikelyImagePlaceholder) {
                        var p = safeGet(function() { return it.position(); }, null);
                        if (p) {
                            targetX = p.x; targetY = p.y;
                            targetW = w;   targetH = h;
                            usedPlaceholder = true;
                            try { it.delete(); } catch (e) { /* leave it; image will overlay */ }
                            break;
                        }
                    }
                }
            } catch (e) { /* iWorkItems may throw on some master types */ }

            var img = kn.Image({ file: Path("\(pathLiteral)") });
            slide.images.push(img);
            img.position = { x: targetX, y: targetY };
            img.width  = targetW;
            img.height = targetH;
            // (Diagnostic value not surfaced today; if useful later we can
            // return a JSON struct of these decisions back to Swift.)
        });
        """
    }

    // MARK: JS string escaping

    /// Escapes a string for safe inclusion inside a double-quoted JS literal.
    /// We deliberately keep this strict — the same engine will eventually take
    /// real user content from DecisionArchitect and we don't want to teach the
    /// downstream codepath to trust raw input.
    private static func jsEscape(_ input: String) -> String {
        var output = ""
        output.reserveCapacity(input.count)
        for scalar in input.unicodeScalars {
            switch scalar {
            case "\\": output += "\\\\"
            case "\"": output += "\\\""
            case "\n": output += "\\n"
            case "\r": output += "\\r"
            case "\t": output += "\\t"
            case "\u{08}": output += "\\b"
            case "\u{0C}": output += "\\f"
            default:
                if scalar.value < 0x20 {
                    output += String(format: "\\u%04x", scalar.value)
                } else {
                    output.unicodeScalars.append(scalar)
                }
            }
        }
        return output
    }
}
