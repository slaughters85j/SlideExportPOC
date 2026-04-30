//
//  ChartRenderer.swift
//  SlideExportPOC
//
//  Renders a SwiftUI ChartView offscreen at slide resolution and writes the
//  result to a PNG file in the temporary directory. The slide engine
//  references that file path inside the JXA script.
//

import SwiftUI
import AppKit

// MARK: - ChartRenderer

enum ChartRenderer {

    enum RenderError: LocalizedError {
        case failedToProduceImage
        case failedToEncodePNG
        case failedToWriteFile(underlying: Error)

        var errorDescription: String? {
            switch self {
            case .failedToProduceImage:
                return "Could not render the chart to an image."
            case .failedToEncodePNG:
                return "Could not encode the chart image as PNG."
            case .failedToWriteFile(let underlying):
                return "Could not write the chart PNG to disk: \(underlying.localizedDescription)"
            }
        }
    }

    // MARK: API

    /// Renders the chart to a 1600×800 PNG (matching the slide image frame
    /// the JXA script positions) and returns the temp file URL.
    @MainActor
    static func renderChartToPNG() throws -> URL {
        let renderSize = CGSize(width: 1600, height: 800)

        let view = ChartView()
            .frame(width: renderSize.width, height: renderSize.height)

        let renderer = ImageRenderer(content: view)
        renderer.proposedSize = ProposedViewSize(width: renderSize.width, height: renderSize.height)
        renderer.scale = 2.0

        guard let nsImage = renderer.nsImage else {
            throw RenderError.failedToProduceImage
        }

        guard
            let tiff = nsImage.tiffRepresentation,
            let bitmap = NSBitmapImageRep(data: tiff),
            let png = bitmap.representation(using: .png, properties: [:])
        else {
            throw RenderError.failedToEncodePNG
        }

        let url = FileManager.default
            .temporaryDirectory
            .appendingPathComponent("SlideExportPOC-chart-\(UUID().uuidString).png")

        do {
            try png.write(to: url, options: .atomic)
        } catch {
            throw RenderError.failedToWriteFile(underlying: error)
        }

        return url
    }
}
