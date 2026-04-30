//
//  SlideItem.swift
//  SlideExportPOC
//
//  Identifies the kinds of slides the user can opt into for export.
//

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
