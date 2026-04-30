//
//  Constants.swift
//  SlideExportPOC
//
//  Hardcoded sample content the PoC exports. Swap these out to test with
//  different shapes of data.
//

import Foundation

// MARK: - Sample Content

enum SampleContent {

    // MARK: Text

    static let title: String = "Quarterly Overview"

    static let body: String = """
    Lorem ipsum dolor sit amet, consectetur adipiscing elit. Sed do eiusmod \
    tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim \
    veniam, quis nostrud exercitation ullamco laboris.
    """

    // MARK: Chart

    struct ChartDatum: Identifiable, Hashable {
        let id = UUID()
        let month: String
        let value: Double
    }

    static let chartData: [ChartDatum] = [
        .init(month: "Jan", value: 42),
        .init(month: "Feb", value: 78),
        .init(month: "Mar", value: 55),
        .init(month: "Apr", value: 90),
        .init(month: "May", value: 63),
    ]
}
