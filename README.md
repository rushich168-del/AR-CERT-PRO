# AR-CERT PRO

## Project Description

AR-CERT PRO is a Microsoft Word VBA application designed to generate professional certificates automatically from Excel data. The project is structured as a commercial-quality automation solution for AaryaRushi Automation Labs, with a clean module layout, reusable components, and a roadmap for controlled feature development.

Current Version: **v0.3**

## Features

- Word-based certificate generation workflow.
- Excel-driven recipient data source.
- Late-bound Excel reader engine.
- Template-based certificate design support.
- Placeholder replacement architecture.
- PDF export module boundary.
- Configuration constants and relative project path helpers.
- Timestamped logging under `Output/Logs`.
- Shared file, folder, filename, and timestamp utilities.
- Organized project structure ready for source control and release packaging.

> Note: v0.3 implements the Excel reader engine only. Word automation, placeholder replacement, PDF generation, mail merge, and GUI work remain planned for future versions.

## Folder Structure

```text
AR-CERT-PRO/
|-- Documentation/
|-- Source/
|   |-- Main.bas
|   |-- Config.bas
|   |-- ExcelReader.bas
|   |-- PlaceholderEngine.bas
|   |-- PDFExporter.bas
|   |-- Utilities.bas
|   `-- Logger.bas
|-- Templates/
|-- Excel/
|-- Output/
|-- Tests/
|-- Releases/
|-- README.md
|-- CHANGELOG.md
`-- LICENSE
```

## Development Roadmap

### v0.1 - Project Foundation - Complete

- Create GitHub-ready folder structure.
- Add documentation files.
- Add VBA source module placeholders.

### v0.2 - Configuration and Logging - Complete

- Implement project configuration constants.
- Add structured logging utilities.
- Define standard error handling patterns.

### v0.3 ✅ Excel Reader Engine

- Initialize and close Excel using late binding.
- Open configured workbooks safely.
- Select worksheets and read cells/ranges.

### v0.4 - Placeholder Engine - Planned

- Detect placeholders in Word templates.
- Replace placeholders with Excel row values.
- Support reusable placeholder naming conventions.

### v0.5 - Certificate Generation Workflow - Planned

- Generate certificates from Word templates.
- Save generated Word documents.
- Add batch processing flow from Excel rows.

### v0.6 - PDF Export and Output Packaging - Planned

- Export generated certificates as PDF.
- Standardize output naming.
- Prepare release-ready output folders.

### v1.0 - Production Release - Planned

- Complete testing.
- Add user documentation.
- Package stable release artifacts.

## Version History

| Version | Date | Summary |
| --- | --- | --- |
| v0.3 | 2026-06-30 | Added late-bound Excel reader engine with workbook, worksheet, cell, and range access. |
| v0.2 | 2026-06-30 | Added configuration constants, relative path helpers, logging, and shared utilities. |
| v0.1 | 2026-06-29 | Created project structure, documentation, and VBA module placeholders. |
