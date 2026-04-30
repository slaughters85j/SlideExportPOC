//
//  SlideItem.swift
//  SlideExportPOC
//
//  Identifies the kinds of slides the user can opt into for export.
//

import CoreGraphics
import Foundation

// MARK: - SlideItem

enum SlideItem: String, CaseIterable, Identifiable, Hashable {
    case titleAndBody
    case chart

    var id: String { rawValue }

    var displayLabel: String {
        switch self {
        case .titleAndBody: return "Title & Body Text"
        case .chart:        return "Bar Chart"
        }
    }
}

// MARK: - SelectedTemplate

/// Identifies a custom template the user has chosen for export. Two paths:
/// - `.file(url)` — the user picked a `.key` or `.pptx` file from disk.
///   Keynote opens it directly via `kn.open(Path(...))` and we use the
///   resulting document as the base.
/// - `.installedTheme(name)` — the user picked a theme from the list of
///   themes already installed in Keynote. We construct a fresh document
///   via `kn.Document({documentTheme: kn.themes["name"]})` — no file open,
///   no install dialog.
///
/// `.kth` files are NOT routed through `.file(...)` because `kn.open` of a
/// `.kth` triggers Keynote's "Add to Theme Chooser" install dialog every
/// time. The intended workflow is: user installs the `.kth` once via
/// Keynote (double-click → Add to Theme Chooser), then picks it from the
/// installed-theme list here.
enum SelectedTemplate: Hashable {
    case file(URL)
    case installedTheme(name: String)

    /// Pure-XML PowerPoint pipeline: skip Keynote entirely, copy the .potx
    /// as the base presentation, inject content into its existing
    /// placeholders, and re-zip as .pptx. See `PPTXTemplateMerger`.
    case potxOverlay(URL)

    var displayLabel: String {
        switch self {
        case .file(let url):           return url.lastPathComponent
        case .installedTheme(let n):   return "Theme: \(n)"
        case .potxOverlay(let url):    return "PowerPoint template: \(url.lastPathComponent)"
        }
    }

    /// Short tag used in JXA diagnostics so the JSON output indicates which
    /// template path drove the document construction.
    var diagnosticMode: String {
        switch self {
        case .file(let url):           return "file:\(url.pathExtension.lowercased())"
        case .installedTheme:          return "installedTheme"
        case .potxOverlay:             return "potxOverlay"
        }
    }
}

// MARK: - SlideAspect

/// Slide dimensions / aspect ratio. Keynote refers to these as "Wide" and
/// "Standard" in its own UI; we use the same labels.
enum SlideAspect: String, CaseIterable, Identifiable, Hashable {
    case wide      // 16:9
    case standard  // 4:3

    var id: String { rawValue }

    var displayLabel: String {
        switch self {
        case .wide:     return "Wide (16:9)"
        case .standard: return "Standard (4:3)"
        }
    }

    /// Slide dimensions, in slide points. These match Keynote's built-in
    /// preset sizes for the Wide and Standard variants of any theme.
    var size: CGSize {
        switch self {
        case .wide:     return CGSize(width: 1920, height: 1080)
        case .standard: return CGSize(width: 1024, height: 768)
        }
    }

    /// Frame for the chart-image region on a slide of this aspect, in slide
    /// points. Defined as a relative inset so 16:9 and 4:3 stay visually
    /// consistent (≈8% horizontal margin, ≈13% vertical margin, the rest is
    /// chart). The chart renderer also uses this size so the rasterized PNG
    /// matches the on-slide frame and stays crisp.
    var chartFrame: CGRect {
        let s = size
        let xMargin = s.width  * 0.083
        let yMargin = s.height * 0.130
        return CGRect(
            x: xMargin,
            y: yMargin,
            width:  s.width  - 2 * xMargin,
            height: s.height - 2 * yMargin
        )
    }
}

// MARK: - ExportFormat

enum ExportFormat: String, CaseIterable, Identifiable, Hashable {
    case keynote
    case powerPoint

    var id: String { rawValue }

    var displayLabel: String {
        switch self {
        case .keynote:     return "Keynote (.key)"
        case .powerPoint:  return "PowerPoint (.pptx)"
        }
    }

    var fileExtension: String {
        switch self {
        case .keynote:     return "key"
        case .powerPoint:  return "pptx"
        }
    }

    /// Format string passed to Keynote's `export ... as: ...` JXA command.
    var keynoteExportFormatString: String {
        switch self {
        case .keynote:     return "Keynote"
        case .powerPoint:  return "Microsoft PowerPoint"
        }
    }
}
