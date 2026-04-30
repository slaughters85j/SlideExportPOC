//
//  ChartView.swift
//  SlideExportPOC
//
//  Reusable Swift Charts bar chart. Used both on-screen (constrained to the
//  left content panel) and offscreen by ChartRenderer when producing the PNG
//  bound for the slide.
//

import SwiftUI
import Charts

// MARK: - ChartView

struct ChartView: View {
    let data: [SampleContent.ChartDatum]

    init(data: [SampleContent.ChartDatum] = SampleContent.chartData) {
        self.data = data
    }

    var body: some View {
        Chart(data) { datum in
            BarMark(
                x: .value("Month", datum.month),
                y: .value("Value", datum.value)
            )
            .foregroundStyle(.blue.gradient)
            .annotation(position: .top, alignment: .center) {
                Text("\(Int(datum.value))")
                    .font(.caption2)
                    .foregroundStyle(.secondary)
            }
        }
        .chartXAxisLabel("Month", position: .bottom, alignment: .center)
        .chartYAxisLabel("Value", position: .leading, alignment: .center)
        .chartYScale(domain: 0...100)
        .padding()
        .background(.background)
    }
}

// MARK: - Preview

#Preview {
    ChartView()
        .frame(width: 400, height: 250)
}
