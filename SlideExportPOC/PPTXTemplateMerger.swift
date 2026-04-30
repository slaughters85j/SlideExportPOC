//
//  PPTXTemplateMerger.swift
//  SlideExportPOC
//
//  Pure-Swift OOXML pipeline: takes a PowerPoint .potx template the user
//  authored (with empty title / body placeholders on its slides) and
//  injects our content into it, producing a real .pptx output.
//
//  Why this exists: Keynote's scripting interface has no path to consume
//  .potx files. Driving Keynote → .pptx then post-processing the XML is the
//  only "polished, deterministic, no manual edits" path for PowerPoint.
//  This file is a working PoC of that path. It deliberately doesn't try to
//  cover every .potx layout the user might author — it covers the specific
//  case where slide 1 has a `<p:ph type="title"/>` placeholder and slide 2
//  has a `<p:ph type="body"/>` placeholder, which is the layout the
//  PoC's authored sample template uses.
//
//  Findings worth knowing for the DA migration are documented inline as
//  comments on each step, plus summarized in the project README.
//

import Foundation

// MARK: - Errors

enum PPTXTemplateMergerError: LocalizedError {
    case templateNotFound(URL)
    case templateUnreadable(URL)
    case unzipFailed(status: Int32)
    case zipFailed(status: Int32)
    case xmlParseFailed(file: String, underlying: Error)
    case xmlSerializeFailed(file: String)
    case missingTitlePlaceholder
    case missingBodyPlaceholder
    case noSlidesInTemplate
    case slideLayoutLookupFailed(reason: String)

    var errorDescription: String? {
        switch self {
        case .templateNotFound(let url):
            return "PowerPoint template not found at \(url.path)."
        case .templateUnreadable(let url):
            return "Could not read \(url.lastPathComponent) — is it a valid .potx zip?"
        case .unzipFailed(let status):
            return "Failed to unzip the .potx (exit status \(status))."
        case .zipFailed(let status):
            return "Failed to re-zip the merged .pptx (exit status \(status))."
        case .xmlParseFailed(let file, let underlying):
            return "Failed to parse \(file): \(underlying.localizedDescription)"
        case .xmlSerializeFailed(let file):
            return "Failed to serialize \(file) back to XML."
        case .missingTitlePlaceholder:
            return "Could not find a title placeholder (<p:ph type=\"title\"/>) on slide 1 of the template."
        case .missingBodyPlaceholder:
            return "Could not find a body placeholder (<p:ph type=\"body\"/>) on slide 2 of the template."
        case .noSlidesInTemplate:
            return "The template has no slides — at least one slide is required."
        case .slideLayoutLookupFailed(let reason):
            return "Could not resolve the body placeholder's geometry from the slide layout: \(reason)"
        }
    }
}

// MARK: - Merger

enum PPTXTemplateMerger {

    /// Diagnostic information returned to the caller about what the merge
    /// did. Useful for surfacing in the host app's success alert / console.
    struct MergeReport {
        let titleInjectedOnSlide: Int?       // 1-based, or nil if no title placeholder was found
        let bodyImageInsertedOnSlide: Int?   // 1-based, or nil if no body placeholder
        let bodyPlaceholderGeometry: String? // a human-readable rect string for diagnostics
        let totalSlidesInOutput: Int
    }

    // MARK: Public entry point

    /// Produces a merged .pptx at `outputURL` by:
    ///   1. Copying the .potx as the base (it's structurally a .pptx with a
    ///      different package MIME type).
    ///   2. Flipping the `[Content_Types].xml` MIME for `/ppt/presentation.xml`
    ///      from `presentationml.template.main+xml` →
    ///      `presentationml.presentation.main+xml`. Without this PowerPoint
    ///      will treat the file as a template.
    ///   3. Injecting `title` into the first slide's title placeholder.
    ///   4. Embedding `chartImage` into the second slide's body placeholder
    ///      area: copies the PNG into `ppt/media/`, registers a relationship
    ///      in `ppt/slides/_rels/slide2.xml.rels`, replaces the body shape
    ///      with a `<p:pic>` element at the body placeholder's geometry
    ///      (resolved from the slide layout).
    ///   5. Re-zipping into a valid .pptx.
    ///
    /// Returns a `MergeReport` describing what was actually changed.
    nonisolated static func merge(
        templateURL: URL,
        title: String,
        chartImageURL: URL?,
        outputURL: URL
    ) throws -> MergeReport {

        guard FileManager.default.fileExists(atPath: templateURL.path) else {
            throw PPTXTemplateMergerError.templateNotFound(templateURL)
        }

        let fm = FileManager.default
        let workspace = fm.temporaryDirectory.appendingPathComponent("PPTXMerge-\(UUID().uuidString)", isDirectory: true)
        defer { try? fm.removeItem(at: workspace) }
        try fm.createDirectory(at: workspace, withIntermediateDirectories: true)

        // 1. Unzip the .potx into the workspace.
        try unzip(archive: templateURL, to: workspace)

        // 2. Flip the package MIME type so PowerPoint treats this as a
        //    .pptx (not a .potx that prompts to "Save As").
        try rewritePackageContentTypes(in: workspace)

        // 3. Enumerate slides, sorted by numeric suffix, so the user's
        //    "slide 1" / "slide 2" mental model maps correctly.
        let slidesDir = workspace.appendingPathComponent("ppt/slides", isDirectory: true)
        let slideFiles = try sortedSlideFiles(in: slidesDir)

        guard !slideFiles.isEmpty else {
            throw PPTXTemplateMergerError.noSlidesInTemplate
        }

        // 4. Inject title text into slide 1's <p:ph type="title"/>.
        let titleSlideURL = slideFiles[0]
        let titleInjected = try injectTitle(intoSlideAt: titleSlideURL, title: title)

        // 5. If we have a chart image and a second slide exists, inject the
        //    image into slide 2's body placeholder area.
        var imageInsertedOnSlide: Int? = nil
        var bodyGeometry: String? = nil
        if let chartImageURL, slideFiles.count >= 2 {
            let chartSlideURL = slideFiles[1]
            let result = try insertChartImage(
                intoSlideAt: chartSlideURL,
                workspace: workspace,
                chartImageURL: chartImageURL
            )
            imageInsertedOnSlide = result.injected ? 2 : nil
            bodyGeometry = result.geometry
        }

        // 6. Re-zip the workspace into the output .pptx.
        if fm.fileExists(atPath: outputURL.path) {
            try fm.removeItem(at: outputURL)
        }
        try zip(directory: workspace, into: outputURL)

        return MergeReport(
            titleInjectedOnSlide: titleInjected ? 1 : nil,
            bodyImageInsertedOnSlide: imageInsertedOnSlide,
            bodyPlaceholderGeometry: bodyGeometry,
            totalSlidesInOutput: slideFiles.count
        )
    }

    // MARK: Step 2 — Flip package content-type

    private nonisolated static func rewritePackageContentTypes(in workspace: URL) throws {
        let url = workspace.appendingPathComponent("[Content_Types].xml")
        var data = try Data(contentsOf: url)
        guard var xml = String(data: data, encoding: .utf8) else {
            throw PPTXTemplateMergerError.xmlParseFailed(file: "[Content_Types].xml",
                underlying: NSError(domain: "PPTX", code: 0, userInfo: [NSLocalizedDescriptionKey: "non-UTF8 content"]))
        }
        let templateMIME = "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
        let presentationMIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
        xml = xml.replacingOccurrences(of: templateMIME, with: presentationMIME)
        data = Data(xml.utf8)
        try data.write(to: url, options: .atomic)
    }

    // MARK: Step 3 — Slide enumeration

    private nonisolated static func sortedSlideFiles(in dir: URL) throws -> [URL] {
        let fm = FileManager.default
        guard fm.fileExists(atPath: dir.path) else { return [] }
        let contents = try fm.contentsOfDirectory(at: dir, includingPropertiesForKeys: nil, options: [.skipsHiddenFiles])
        let slides = contents.filter {
            $0.lastPathComponent.hasPrefix("slide") && $0.pathExtension.lowercased() == "xml"
        }
        // Sort by trailing numeric: slide1.xml, slide2.xml, ..., slide10.xml.
        return slides.sorted { lhs, rhs in
            slideIndex(lhs) < slideIndex(rhs)
        }
    }

    private nonisolated static func slideIndex(_ url: URL) -> Int {
        let stem = url.deletingPathExtension().lastPathComponent
        let digits = stem.drop(while: { !$0.isNumber })
        return Int(digits) ?? Int.max
    }

    // MARK: Step 4 — Inject title text

    /// Finds the first `<p:sp>` containing `<p:ph type="title"/>` (or
    /// `ctrTitle`) and replaces its text with `title`. Returns true on hit.
    private nonisolated static func injectTitle(intoSlideAt url: URL, title: String) throws -> Bool {
        let doc = try parseXML(at: url)
        let root = doc.rootElement()

        // Find the title-typed placeholder shape.
        guard let titleShape = findPlaceholderShape(in: root, ofTypes: ["title", "ctrTitle"]) else {
            return false
        }

        // The shape's text is in <p:txBody><a:p><a:r><a:t>…</a:t></a:r></a:p>.
        // The .potx may have an empty <a:p><a:endParaRPr/></a:p> instead.
        // Either way we replace the entire <p:txBody> contents with one
        // paragraph containing one run with our text — that's the canonical
        // form PowerPoint emits and round-trips cleanly.
        try replaceTextBody(of: titleShape, with: title)
        try writeXML(doc: doc, to: url)
        return true
    }

    // MARK: Step 5 — Insert chart image into body placeholder

    private nonisolated static func insertChartImage(
        intoSlideAt slideURL: URL,
        workspace: URL,
        chartImageURL: URL
    ) throws -> (injected: Bool, geometry: String?) {

        let doc = try parseXML(at: slideURL)
        let root = doc.rootElement()

        guard let bodyShape = findPlaceholderShape(in: root, ofTypes: ["body"]) else {
            return (false, nil)
        }

        // Resolve the body's geometry. Three-tier fallback:
        //
        //   1. Layout's body placeholder <a:xfrm>. Many templates inherit
        //      from the master rather than declaring xfrm at the layout
        //      level, so this is often nil. (Verified: the user's sample
        //      .potx routes slide 2 through slideLayout18 whose body
        //      placeholder has no xfrm.)
        //   2. Slide-relative frame derived from <p:sldSz> in
        //      ppt/presentation.xml. Always works as long as the package
        //      is well-formed; produces a centered chart at consistent 8%
        //      horizontal / 13% vertical margins regardless of canvas size.
        //   3. Hardcoded sensible defaults (last resort).
        let layoutGeometry = try resolveBodyGeometry(slideURL: slideURL, workspace: workspace)
        let geometry: EmuRect? = layoutGeometry ?? resolveSlideRelativeChartFrame(in: workspace)

        // Add the image to ppt/media/, register a relationship for it in
        // the slide's _rels file, and replace the body shape with a <p:pic>
        // sized to the body's geometry (or sensibly defaulted).
        let mediaName = "image-poc-\(UUID().uuidString).png"
        let mediaURL = workspace.appendingPathComponent("ppt/media/\(mediaName)")
        try FileManager.default.createDirectory(
            at: mediaURL.deletingLastPathComponent(),
            withIntermediateDirectories: true
        )
        try FileManager.default.copyItem(at: chartImageURL, to: mediaURL)

        let rId = try registerRelationship(
            slideURL: slideURL,
            mediaName: mediaName
        )

        guard let parent = bodyShape.parent as? XMLElement else {
            return (false, geometry?.debugDescription)
        }

        // Build the <p:pic> element to replace the body <p:sp>.
        let pic = makePicElement(rId: rId, frame: geometry)
        let bodyIndex = parent.children?.firstIndex(where: { $0 === bodyShape }) ?? 0
        parent.removeChild(at: bodyIndex)
        parent.insertChild(pic, at: bodyIndex)

        try writeXML(doc: doc, to: slideURL)

        return (true, geometry?.debugDescription)
    }

    // MARK: Body geometry resolution from layout

    private struct EmuRect: CustomDebugStringConvertible, Sendable {
        let x: Int64
        let y: Int64
        let cx: Int64
        let cy: Int64
        nonisolated var debugDescription: String { "x=\(x) y=\(y) cx=\(cx) cy=\(cy) (EMUs)" }
    }

    /// Reads `<p:sldSz cx="..." cy="..."/>` from `ppt/presentation.xml` and
    /// returns a chart frame inset by 8% horizontally / 13% vertically.
    /// Returns nil if presentation.xml or the size element is missing.
    private nonisolated static func resolveSlideRelativeChartFrame(in workspace: URL) -> EmuRect? {
        let presURL = workspace.appendingPathComponent("ppt/presentation.xml")
        guard
            let data = try? Data(contentsOf: presURL),
            let presDoc = try? XMLDocument(data: data, options: [.nodePreserveAll]),
            let root = presDoc.rootElement(),
            let sldSz = firstDescendant(of: root, named: "p:sldSz")
        else {
            return nil
        }
        guard
            let cxStr = sldSz.attribute(forName: "cx")?.stringValue,
            let cyStr = sldSz.attribute(forName: "cy")?.stringValue,
            let cx = Int64(cxStr),
            let cy = Int64(cyStr),
            cx > 0, cy > 0
        else {
            return nil
        }
        let xMargin = Int64(Double(cx) * 0.083)
        let yMargin = Int64(Double(cy) * 0.130)
        return EmuRect(
            x: xMargin,
            y: yMargin,
            cx: cx - 2 * xMargin,
            cy: cy - 2 * yMargin
        )
    }

    private nonisolated static func resolveBodyGeometry(slideURL: URL, workspace: URL) throws -> EmuRect? {
        // 1. Find slide.xml.rels for this slide.
        let slidesRelsDir = slideURL.deletingLastPathComponent().appendingPathComponent("_rels", isDirectory: true)
        let relsFile = slidesRelsDir.appendingPathComponent(slideURL.lastPathComponent + ".rels")
        guard FileManager.default.fileExists(atPath: relsFile.path) else { return nil }

        let relsDoc = try parseXML(at: relsFile)
        guard let relsRoot = relsDoc.rootElement() else { return nil }

        // 2. Locate the relationship whose Type is slideLayout.
        let layoutRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
        var layoutTarget: String?
        for child in relsRoot.children ?? [] {
            guard let elem = child as? XMLElement else { continue }
            let type = elem.attribute(forName: "Type")?.stringValue ?? ""
            if type == layoutRelType {
                layoutTarget = elem.attribute(forName: "Target")?.stringValue
                break
            }
        }
        guard let layoutTarget else { return nil }

        // 3. Resolve the layout path. Targets are usually "../slideLayouts/slideLayoutN.xml".
        let layoutURL = slideURL.deletingLastPathComponent().appendingPathComponent(layoutTarget).standardizedFileURL
        guard FileManager.default.fileExists(atPath: layoutURL.path) else { return nil }

        let layoutDoc = try parseXML(at: layoutURL)
        guard let layoutRoot = layoutDoc.rootElement() else { return nil }

        // 4. Find the body placeholder in the layout, then its <a:xfrm>.
        guard let bodyShape = findPlaceholderShape(in: layoutRoot, ofTypes: ["body"]) else { return nil }
        guard let xfrm = firstDescendant(of: bodyShape, named: "a:xfrm") else { return nil }
        guard
            let off = firstDescendant(of: xfrm, named: "a:off"),
            let ext = firstDescendant(of: xfrm, named: "a:ext")
        else { return nil }

        let x = Int64(off.attribute(forName: "x")?.stringValue ?? "0") ?? 0
        let y = Int64(off.attribute(forName: "y")?.stringValue ?? "0") ?? 0
        let cx = Int64(ext.attribute(forName: "cx")?.stringValue ?? "0") ?? 0
        let cy = Int64(ext.attribute(forName: "cy")?.stringValue ?? "0") ?? 0
        guard cx > 0 && cy > 0 else { return nil }

        return EmuRect(x: x, y: y, cx: cx, cy: cy)
    }

    // MARK: Slide rels — register image relationship

    /// Adds a Relationship of type "image" pointing to "../media/<name>".
    /// Picks a fresh rId not already in use. Returns the rId.
    private nonisolated static func registerRelationship(slideURL: URL, mediaName: String) throws -> String {
        let slidesRelsDir = slideURL.deletingLastPathComponent().appendingPathComponent("_rels", isDirectory: true)
        try FileManager.default.createDirectory(at: slidesRelsDir, withIntermediateDirectories: true)
        let relsFile = slidesRelsDir.appendingPathComponent(slideURL.lastPathComponent + ".rels")

        let imageRelType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        let target = "../media/\(mediaName)"

        let doc: XMLDocument
        let root: XMLElement
        if FileManager.default.fileExists(atPath: relsFile.path) {
            doc = try parseXML(at: relsFile)
            guard let r = doc.rootElement() else {
                throw PPTXTemplateMergerError.xmlParseFailed(file: relsFile.lastPathComponent,
                    underlying: NSError(domain: "PPTX", code: 0))
            }
            root = r
        } else {
            doc = XMLDocument(rootElement: nil)
            doc.version = "1.0"
            doc.characterEncoding = "UTF-8"
            doc.isStandalone = true
            let r = XMLElement(name: "Relationships")
            r.addAttribute(XMLNode.attribute(withName: "xmlns", stringValue: "http://schemas.openxmlformats.org/package/2006/relationships") as! XMLNode)
            doc.setRootElement(r)
            root = r
        }

        // Find max existing rIdN.
        var maxNum = 0
        for child in root.children ?? [] {
            guard let elem = child as? XMLElement else { continue }
            if let id = elem.attribute(forName: "Id")?.stringValue, id.hasPrefix("rId"),
               let n = Int(id.dropFirst(3)) {
                maxNum = max(maxNum, n)
            }
        }
        let rId = "rId\(maxNum + 1)"

        let rel = XMLElement(name: "Relationship")
        rel.addAttribute(XMLNode.attribute(withName: "Id",     stringValue: rId)             as! XMLNode)
        rel.addAttribute(XMLNode.attribute(withName: "Type",   stringValue: imageRelType)    as! XMLNode)
        rel.addAttribute(XMLNode.attribute(withName: "Target", stringValue: target)          as! XMLNode)
        root.addChild(rel)

        try writeXML(doc: doc, to: relsFile)
        return rId
    }

    // MARK: Build a <p:pic> element

    private nonisolated static func makePicElement(rId: String, frame: EmuRect?) -> XMLElement {
        // Geometry: prefer the body placeholder's frame; otherwise pick a
        // reasonable default (centered, mid-size). EMUs: 914400 = 1 inch.
        let off = frame.map { (x: $0.x, y: $0.y) } ?? (x: 1_524_000, y: 1_524_000)   // ~1.67"
        let ext = frame.map { (cx: $0.cx, cy: $0.cy) } ?? (cx: 6_096_000, cy: 4_572_000) // 6.67" x 5.0"

        // The XML below is the canonical form PowerPoint emits for an
        // embedded picture. Built via XMLElement so attribute escaping is
        // correct; the embed reference goes via the rId we just created.
        let pic = XMLElement(name: "p:pic")

        let nvPicPr = XMLElement(name: "p:nvPicPr")
        let cNvPr = XMLElement(name: "p:cNvPr")
        cNvPr.addAttribute(XMLNode.attribute(withName: "id", stringValue: "1000") as! XMLNode)
        cNvPr.addAttribute(XMLNode.attribute(withName: "name", stringValue: "ChartImage") as! XMLNode)
        nvPicPr.addChild(cNvPr)
        let cNvPicPr = XMLElement(name: "p:cNvPicPr")
        let picLocks = XMLElement(name: "a:picLocks")
        picLocks.addAttribute(XMLNode.attribute(withName: "noChangeAspect", stringValue: "1") as! XMLNode)
        cNvPicPr.addChild(picLocks)
        nvPicPr.addChild(cNvPicPr)
        nvPicPr.addChild(XMLElement(name: "p:nvPr"))
        pic.addChild(nvPicPr)

        let blipFill = XMLElement(name: "p:blipFill")
        let blip = XMLElement(name: "a:blip")
        blip.addAttribute(XMLNode.attribute(withName: "r:embed", stringValue: rId) as! XMLNode)
        blipFill.addChild(blip)
        let stretch = XMLElement(name: "a:stretch")
        stretch.addChild(XMLElement(name: "a:fillRect"))
        blipFill.addChild(stretch)
        pic.addChild(blipFill)

        let spPr = XMLElement(name: "p:spPr")
        let xfrm = XMLElement(name: "a:xfrm")
        let offEl = XMLElement(name: "a:off")
        offEl.addAttribute(XMLNode.attribute(withName: "x", stringValue: "\(off.x)") as! XMLNode)
        offEl.addAttribute(XMLNode.attribute(withName: "y", stringValue: "\(off.y)") as! XMLNode)
        xfrm.addChild(offEl)
        let extEl = XMLElement(name: "a:ext")
        extEl.addAttribute(XMLNode.attribute(withName: "cx", stringValue: "\(ext.cx)") as! XMLNode)
        extEl.addAttribute(XMLNode.attribute(withName: "cy", stringValue: "\(ext.cy)") as! XMLNode)
        xfrm.addChild(extEl)
        spPr.addChild(xfrm)
        let prstGeom = XMLElement(name: "a:prstGeom")
        prstGeom.addAttribute(XMLNode.attribute(withName: "prst", stringValue: "rect") as! XMLNode)
        prstGeom.addChild(XMLElement(name: "a:avLst"))
        spPr.addChild(prstGeom)
        pic.addChild(spPr)

        return pic
    }

    // MARK: Placeholder lookup helpers

    /// Returns the first `<p:sp>` element under `root` whose `<p:nvSpPr>`
    /// contains `<p:nvPr><p:ph type="X"/></p:nvPr>` for some X in `types`.
    private nonisolated static func findPlaceholderShape(in root: XMLElement?, ofTypes types: [String]) -> XMLElement? {
        guard let root else { return nil }
        let shapes = descendants(of: root, named: "p:sp")
        for shape in shapes {
            guard let ph = firstDescendant(of: shape, named: "p:ph") else { continue }
            let phType = ph.attribute(forName: "type")?.stringValue ?? "body" // OOXML default for missing type
            if types.contains(phType) {
                return shape
            }
        }
        return nil
    }

    private nonisolated static func descendants(of element: XMLElement, named name: String) -> [XMLElement] {
        var results: [XMLElement] = []
        for child in element.children ?? [] {
            guard let elem = child as? XMLElement else { continue }
            if elem.name == name { results.append(elem) }
            results.append(contentsOf: descendants(of: elem, named: name))
        }
        return results
    }

    private nonisolated static func firstDescendant(of element: XMLElement, named name: String) -> XMLElement? {
        for child in element.children ?? [] {
            guard let elem = child as? XMLElement else { continue }
            if elem.name == name { return elem }
            if let nested = firstDescendant(of: elem, named: name) { return nested }
        }
        return nil
    }

    // MARK: Replace text body content

    /// Replaces the contents of a `<p:sp>`'s `<p:txBody>` with one paragraph
    /// holding one run carrying `text`. The empty placeholder pattern in
    /// the user's .potx is `<a:p><a:endParaRPr/></a:p>` — replacing the
    /// entire txBody with a populated paragraph is the canonical way to
    /// give it actual content.
    private nonisolated static func replaceTextBody(of shape: XMLElement, with text: String) throws {
        guard let txBody = firstDescendant(of: shape, named: "p:txBody") else { return }

        // Preserve any existing <a:bodyPr> and <a:lstStyle> children — those
        // hold formatting hints we don't want to drop. Replace only the
        // <a:p> children.
        var bodyPr: XMLElement?
        var lstStyle: XMLElement?
        for child in txBody.children ?? [] {
            guard let elem = child as? XMLElement else { continue }
            if elem.name == "a:bodyPr"  { bodyPr  = elem.copy() as? XMLElement }
            if elem.name == "a:lstStyle"{ lstStyle = elem.copy() as? XMLElement }
        }
        // Remove all current children.
        while let n = txBody.children?.last {
            txBody.removeChild(at: txBody.childCount - 1)
            _ = n
        }

        if let bodyPr   { txBody.addChild(bodyPr)   } else { txBody.addChild(XMLElement(name: "a:bodyPr")) }
        if let lstStyle { txBody.addChild(lstStyle) } else { txBody.addChild(XMLElement(name: "a:lstStyle")) }

        let p = XMLElement(name: "a:p")
        let r = XMLElement(name: "a:r")
        let rPr = XMLElement(name: "a:rPr")
        rPr.addAttribute(XMLNode.attribute(withName: "lang", stringValue: "en-US") as! XMLNode)
        rPr.addAttribute(XMLNode.attribute(withName: "dirty", stringValue: "0") as! XMLNode)
        let t = XMLElement(name: "a:t", stringValue: text)
        r.addChild(rPr)
        r.addChild(t)
        p.addChild(r)
        txBody.addChild(p)
    }

    // MARK: XML I/O helpers

    private nonisolated static func parseXML(at url: URL) throws -> XMLDocument {
        let data = try Data(contentsOf: url)
        do {
            let doc = try XMLDocument(data: data, options: [.nodePreserveAll])
            return doc
        } catch {
            throw PPTXTemplateMergerError.xmlParseFailed(file: url.lastPathComponent, underlying: error)
        }
    }

    private nonisolated static func writeXML(doc: XMLDocument, to url: URL) throws {
        let data = doc.xmlData(options: [.nodeCompactEmptyElement])
        try data.write(to: url, options: .atomic)
    }

    // MARK: zip / unzip via /usr/bin

    private nonisolated static func unzip(archive: URL, to destination: URL) throws {
        try FileManager.default.createDirectory(at: destination, withIntermediateDirectories: true)
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/unzip")
        process.arguments = ["-q", "-o", archive.path, "-d", destination.path]
        process.standardOutput = Pipe()
        process.standardError = Pipe()
        try process.run()
        process.waitUntilExit()
        guard process.terminationStatus == 0 else {
            throw PPTXTemplateMergerError.unzipFailed(status: process.terminationStatus)
        }
    }

    private nonisolated static func zip(directory: URL, into output: URL) throws {
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/zip")
        // -r recurse, -q quiet, -X strip extra attributes (avoids spurious extra fields).
        process.arguments = ["-r", "-q", "-X", output.path, "."]
        process.currentDirectoryURL = directory
        process.standardOutput = Pipe()
        process.standardError = Pipe()
        try process.run()
        process.waitUntilExit()
        guard process.terminationStatus == 0 else {
            throw PPTXTemplateMergerError.zipFailed(status: process.terminationStatus)
        }
    }
}
