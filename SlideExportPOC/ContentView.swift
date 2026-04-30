//
//  ContentView.swift
//  SlideExportPOC
//
//  Split layout: hardcoded sample content on the left (the things we'd
//  export), export controls on the right.
//

import SwiftUI
import AppKit

// MARK: - ContentView

struct ContentView: View {

    // MARK: State

    @State private var showingExportSheet = false
    @State private var isExporting = false
    @State private var exportResult: ExportResult? = nil

    // MARK: Body

    var body: some View {
        HSplitView {
            contentPanel
                .frame(minWidth: 480, idealWidth: 560)
            exportPanel
                .frame(minWidth: 260, idealWidth: 300)
        }
        .frame(minWidth: 800, minHeight: 600)
        .sheet(isPresented: $showingExportSheet) {
            ExportSelectionView { items, format, template in
                handleExport(items: items, format: format, template: template)
            }
        }
        .alert(
            alertTitle,
            isPresented: alertPresentedBinding,
            presenting: exportResult
        ) { result in
            alertButtons(for: result)
        } message: { result in
            alertMessage(for: result)
        }
        .overlay {
            if isExporting { exportingOverlay }
        }
    }

    // MARK: Alert wiring

    private var alertPresentedBinding: Binding<Bool> {
        Binding(
            get: { exportResult != nil },
            set: { if !$0 { exportResult = nil } }
        )
    }

    private var alertTitle: String {
        switch exportResult {
        case .success:                  return "Exported to Desktop"
        case .permissionDenied:         return "Automation Permission Required"
        case .failure, .none:           return "Export Failed"
        }
    }

    @ViewBuilder
    private func alertButtons(for result: ExportResult) -> some View {
        switch result {
        case .success(let url):
            Button("Reveal in Finder") {
                NSWorkspace.shared.activateFileViewerSelecting([url])
            }
            Button("Done", role: .cancel) { }

        case .permissionDenied:
            Button("Open Automation Settings") {
                openAutomationSettings()
            }
            Button("Reset Permission & Quit") {
                resetAutomationPermissionAndQuit()
            }
            Button("Dismiss", role: .cancel) { }

        case .failure:
            Button("OK", role: .cancel) { }
        }
    }

    @ViewBuilder
    private func alertMessage(for result: ExportResult) -> some View {
        switch result {
        case .success(let url):
            Text(url.lastPathComponent)
        case .permissionDenied:
            Text("""
            macOS blocked SlideExportPOC from controlling Keynote.

            To grant access:
            1. Click “Open Automation Settings”.
            2. Find SlideExportPOC in the list and enable Keynote.
            3. Try the export again.

            If SlideExportPOC isn’t listed, click “Reset Permission & Quit” — that clears the cached denial. Reopen the app and try the export; macOS will prompt for permission.
            """)
        case .failure(let message):
            Text(message)
        }
    }

    // MARK: Permission helpers

    private func openAutomationSettings() {
        // Deep-link directly to Privacy & Security → Automation. The URL
        // scheme is supported on macOS 13+ and macOS 14 redirects the legacy
        // pref pane URL to System Settings.
        if let url = URL(string: "x-apple.systempreferences:com.apple.preference.security?Privacy_Automation") {
            NSWorkspace.shared.open(url)
        }
    }

    private func resetAutomationPermissionAndQuit() {
        guard let bundleID = Bundle.main.bundleIdentifier else { return }
        let task = Process()
        task.executableURL = URL(fileURLWithPath: "/usr/bin/tccutil")
        task.arguments = ["reset", "AppleEvents", bundleID]
        do {
            try task.run()
            task.waitUntilExit()
            #if DEBUG
            print("[ContentView] tccutil reset AppleEvents \(bundleID) → exit \(task.terminationStatus)")
            #endif
        } catch {
            #if DEBUG
            print("[ContentView] tccutil failed to launch: \(error)")
            #endif
        }
        // Quit so the next launch is a fresh process — required for the new
        // permission prompt to appear cleanly.
        NSApp.terminate(nil)
    }

    // MARK: Panels

    private var contentPanel: some View {
        VStack(alignment: .leading, spacing: 20) {
            panelHeader("Content")

            VStack(alignment: .leading, spacing: 8) {
                Text(SampleContent.title)
                    .font(.title2.weight(.semibold))
                Text(SampleContent.body)
                    .font(.body)
                    .foregroundStyle(.secondary)
                    .fixedSize(horizontal: false, vertical: true)
            }

            ChartView()
                .frame(width: 400, height: 250)
                .background(
                    RoundedRectangle(cornerRadius: 8)
                        .stroke(.separator)
                )

            Spacer()
        }
        .padding(20)
        .frame(maxWidth: .infinity, maxHeight: .infinity, alignment: .topLeading)
    }

    private var exportPanel: some View {
        VStack(alignment: .leading, spacing: 20) {
            panelHeader("Export")

            Button {
                showingExportSheet = true
            } label: {
                Label("Export to Slides", systemImage: "square.and.arrow.up")
                    .frame(maxWidth: .infinity)
            }
            .buttonStyle(.borderedProminent)
            .controlSize(.large)
            .disabled(isExporting)

            Text("Builds a Keynote document via OSAKit/JXA and exports to .pptx or .key on your Desktop.")
                .font(.caption)
                .foregroundStyle(.secondary)
                .fixedSize(horizontal: false, vertical: true)

            Spacer()
        }
        .padding(20)
        .frame(maxWidth: .infinity, maxHeight: .infinity, alignment: .topLeading)
    }

    private func panelHeader(_ title: String) -> some View {
        Text(title)
            .font(.headline)
            .foregroundStyle(.secondary)
            .textCase(.uppercase)
            .tracking(0.5)
    }

    private var exportingOverlay: some View {
        ZStack {
            Color.black.opacity(0.25)
            VStack(spacing: 12) {
                ProgressView()
                    .controlSize(.large)
                Text("Exporting…")
                    .font(.callout)
            }
            .padding(24)
            .background(.regularMaterial, in: RoundedRectangle(cornerRadius: 12))
        }
        .ignoresSafeArea()
    }

    // MARK: Export handler

    private func handleExport(items: [SlideItem], format: ExportFormat, template: URL?) {
        isExporting = true
        Task {
            defer { isExporting = false }
            do {
                let url = try await KeynoteExporter.exportToKeynote(
                    items: items,
                    format: format,
                    customTemplateURL: template
                )
                exportResult = .success(url)
            } catch let kerr as KeynoteExporterError {
                if case .automationPermissionDenied = kerr {
                    exportResult = .permissionDenied
                } else {
                    exportResult = .failure(kerr.localizedDescription)
                }
            } catch {
                exportResult = .failure(error.localizedDescription)
            }
        }
    }
}

// MARK: - ExportResult

enum ExportResult: Identifiable {
    case success(URL)
    case permissionDenied
    case failure(String)

    var id: String {
        switch self {
        case .success(let url):       return "success-\(url.path)"
        case .permissionDenied:       return "permission"
        case .failure(let msg):       return "failure-\(msg)"
        }
    }
}

// MARK: - Preview

#Preview {
    ContentView()
}
