//
//  ExportSelectionView.swift
//  SlideExportPOC
//
//  Modal sheet presented from ContentView. Lets the user pick which slide
//  items to include, the output format, and an optional custom template
//  before kicking off the export.
//

import SwiftUI
import AppKit
import UniformTypeIdentifiers

// MARK: - ExportSelectionView

struct ExportSelectionView: View {

    // MARK: Inputs / outputs

    /// Called when the user taps Export. The view dismisses itself first.
    var onExport: (_ items: [SlideItem], _ format: ExportFormat, _ aspect: SlideAspect, _ template: SelectedTemplate?) -> Void

    // MARK: Local state

    @Environment(\.dismiss) private var dismiss

    @State private var selectedItems: Set<SlideItem> = Set(SlideItem.allCases)
    @State private var format: ExportFormat = .powerPoint
    @State private var aspect: SlideAspect = .wide
    @State private var selectedTemplate: SelectedTemplate? = nil

    @State private var installedThemes: [String] = []
    @State private var loadingThemes: Bool = false
    @State private var showingThemePicker: Bool = false

    private let keynoteStatus: KeynoteInstallStatus = KeynoteExporter.keynoteInstallStatus()
    private var isKeynoteInstalled: Bool { keynoteStatus.isUsable }

    // MARK: Body

    var body: some View {
        VStack(alignment: .leading, spacing: 0) {
            header

            Divider()

            ScrollView {
                VStack(alignment: .leading, spacing: 20) {
                    if !isKeynoteInstalled {
                        keynoteMissingNotice
                    }

                    slidesSection
                    formatSection
                    aspectSection
                    customTemplateSection
                }
                .padding(20)
            }

            Divider()

            footer
        }
        .frame(width: 480, height: 480)
    }

    // MARK: Sections

    private var header: some View {
        Text("Select Slides to Export")
            .font(.title3.weight(.semibold))
            .frame(maxWidth: .infinity, alignment: .leading)
            .padding(.horizontal, 20)
            .padding(.vertical, 14)
    }

    private var keynoteMissingNotice: some View {
        VStack(alignment: .leading, spacing: 8) {
            HStack(alignment: .top, spacing: 10) {
                Image(systemName: "exclamationmark.triangle.fill")
                    .foregroundStyle(.orange)
                VStack(alignment: .leading, spacing: 4) {
                    Text("Keynote is required to export — even when exporting to PowerPoint.")
                        .font(.callout.weight(.semibold))
                        .fixedSize(horizontal: false, vertical: true)
                    Text(keynoteStatus.diagnostic)
                        .font(.callout)
                        .foregroundStyle(.secondary)
                        .fixedSize(horizontal: false, vertical: true)
                }
            }
            if let foundURL = keynoteStatus.foundURL, !keynoteStatus.isUsable {
                Text("Detected at: \(foundURL.path)")
                    .font(.caption.monospaced())
                    .foregroundStyle(.secondary)
                    .textSelection(.enabled)
            }
        }
        .padding(12)
        .background(.orange.opacity(0.1), in: RoundedRectangle(cornerRadius: 8))
    }

    private var slidesSection: some View {
        VStack(alignment: .leading, spacing: 8) {
            sectionLabel("Slides")
            ForEach(SlideItem.allCases) { item in
                Toggle(isOn: binding(for: item)) {
                    Text(item.displayLabel)
                }
                .toggleStyle(.checkbox)
                .disabled(!isKeynoteInstalled)
            }
        }
    }

    private var formatSection: some View {
        VStack(alignment: .leading, spacing: 8) {
            sectionLabel("Format")
            Picker("Format", selection: $format) {
                ForEach(ExportFormat.allCases) { fmt in
                    Text(fmt.displayLabel).tag(fmt)
                }
            }
            .pickerStyle(.segmented)
            .labelsHidden()
            .disabled(!isKeynoteInstalled)
        }
    }

    private var aspectSection: some View {
        VStack(alignment: .leading, spacing: 8) {
            sectionLabel("Aspect")
            Picker("Aspect", selection: $aspect) {
                ForEach(SlideAspect.allCases) { a in
                    Text(a.displayLabel).tag(a)
                }
            }
            .pickerStyle(.segmented)
            .labelsHidden()
            .disabled(!isKeynoteInstalled || selectedTemplate != nil)

            if selectedTemplate != nil {
                Text("Aspect is set by the template when one is selected.")
                    .font(.caption)
                    .foregroundStyle(.secondary)
            }
        }
    }

    private var customTemplateSection: some View {
        VStack(alignment: .leading, spacing: 8) {
            sectionLabel("Custom Template (optional)")
            HStack(spacing: 8) {
                Text(selectedTemplate?.displayLabel ?? "No template selected")
                    .font(.callout)
                    .foregroundStyle(selectedTemplate == nil ? .secondary : .primary)
                    .lineLimit(1)
                    .truncationMode(.middle)
                    .frame(maxWidth: .infinity, alignment: .leading)

                if selectedTemplate != nil {
                    Button("Clear") { selectedTemplate = nil }
                        .buttonStyle(.borderless)
                }

                Menu("Choose…") {
                    Button("Installed Keynote Theme…") {
                        Task { await loadAndShowThemes() }
                    }
                    Button("Presentation File (.key / .pptx)…") {
                        chooseTemplateFile()
                    }
                }
                .disabled(!isKeynoteInstalled)
            }
            Text("Tip: for .kth themes, install once via Keynote (double-click → Add to Theme Chooser), then pick from Installed Keynote Theme. .potx PowerPoint templates aren’t supported — convert to .pptx first.")
                .font(.caption)
                .foregroundStyle(.secondary)
                .fixedSize(horizontal: false, vertical: true)
        }
        .sheet(isPresented: $showingThemePicker) {
            installedThemePicker
        }
    }

    // MARK: Installed-theme picker sheet

    private var installedThemePicker: some View {
        VStack(alignment: .leading, spacing: 0) {
            HStack {
                Text("Choose Installed Keynote Theme")
                    .font(.headline)
                Spacer()
                Button("Cancel") { showingThemePicker = false }
                    .keyboardShortcut(.cancelAction)
            }
            .padding(16)

            Divider()

            if loadingThemes {
                ProgressView().padding(40)
                    .frame(maxWidth: .infinity)
            } else if installedThemes.isEmpty {
                Text("No themes found. Install your .kth via Keynote first (double-click the file → Add to Theme Chooser).")
                    .font(.callout)
                    .foregroundStyle(.secondary)
                    .padding(20)
                    .frame(maxWidth: .infinity, alignment: .leading)
            } else {
                List {
                    ForEach(installedThemes, id: \.self) { name in
                        Button {
                            selectedTemplate = .installedTheme(name: name)
                            showingThemePicker = false
                        } label: {
                            HStack {
                                Text(name)
                                Spacer()
                                if case .installedTheme(let n) = selectedTemplate, n == name {
                                    Image(systemName: "checkmark")
                                        .foregroundStyle(.tint)
                                }
                            }
                            .contentShape(Rectangle())
                        }
                        .buttonStyle(.plain)
                    }
                }
            }
        }
        .frame(width: 420, height: 480)
    }

    private func loadAndShowThemes() async {
        loadingThemes = true
        showingThemePicker = true
        do {
            installedThemes = try await KeynoteExporter.listInstalledThemes()
        } catch {
            #if DEBUG
            print("[ExportSelectionView] listInstalledThemes failed: \(error)")
            #endif
            installedThemes = []
        }
        loadingThemes = false
    }

    private var footer: some View {
        HStack {
            Spacer()
            Button("Cancel", role: .cancel) { dismiss() }
                .keyboardShortcut(.cancelAction)
            Button("Export") {
                let items = SlideItem.allCases.filter { selectedItems.contains($0) }
                let template = selectedTemplate
                let chosenFormat = format
                let chosenAspect = aspect
                dismiss()
                onExport(items, chosenFormat, chosenAspect, template)
            }
            .buttonStyle(.borderedProminent)
            .keyboardShortcut(.defaultAction)
            .disabled(!isKeynoteInstalled || selectedItems.isEmpty)
        }
        .padding(.horizontal, 20)
        .padding(.vertical, 14)
    }

    // MARK: Helpers

    private func sectionLabel(_ text: String) -> some View {
        Text(text)
            .font(.subheadline.weight(.semibold))
            .foregroundStyle(.secondary)
    }

    private func binding(for item: SlideItem) -> Binding<Bool> {
        Binding(
            get: { selectedItems.contains(item) },
            set: { isOn in
                if isOn { selectedItems.insert(item) }
                else    { selectedItems.remove(item) }
            }
        )
    }

    private func chooseTemplateFile() {
        let panel = NSOpenPanel()
        panel.title = "Choose a Presentation File"
        panel.canChooseDirectories = false
        panel.canChooseFiles = true
        panel.allowsMultipleSelection = false

        // Only .key and .pptx — .kth flows through the installed-theme path
        // because kn.open(.kth) triggers a modal "Add to Theme Chooser"
        // dialog every time. .potx is not supported by Keynote scripting.
        var allowedTypes: [UTType] = []
        if let key  = UTType(filenameExtension: "key")  { allowedTypes.append(key) }
        if let pptx = UTType(filenameExtension: "pptx") { allowedTypes.append(pptx) }
        if !allowedTypes.isEmpty {
            panel.allowedContentTypes = allowedTypes
        }

        if panel.runModal() == .OK, let url = panel.url {
            selectedTemplate = .file(url)
        }
    }
}

// MARK: - Preview

#Preview {
    ExportSelectionView { _, _, _, _ in }
}
