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
