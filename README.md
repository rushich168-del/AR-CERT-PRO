# AR-CERT PRO

## Project Description

AR-CERT PRO is a Microsoft Word VBA application designed to generate professional certificates automatically from Excel data. The project is structured as a commercial-quality automation solution for AaryaRushi Automation Labs, with a clean module layout, reusable components, and a roadmap for controlled feature development.

Current Version: **v0.1**

## Features

- Word-based certificate generation workflow.
- Excel-driven recipient data source.
- Template-based certificate design support.
- Placeholder replacement architecture.
- PDF export module boundary.
- Logging and utility module boundaries.
- Organized project structure ready for source control and release packaging.

> Note: v0.1 establishes the project structure and documentation only. Business logic will be implemented in future versions.

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

### v0.1 - Project Foundation

- Create GitHub-ready folder structure.
- Add documentation files.
- Add VBA source module placeholders.

### v0.2 - Configuration and Logging

- Implement project configuration constants.
- Add structured logging utilities.
- Define standard error handling patterns.

### v0.3 - Excel Data Reader

- Read recipient data from Excel workbooks.
- Validate required columns.
- Normalize input values for certificate generation.

### v0.4 - Placeholder Engine

- Detect placeholders in Word templates.
- Replace placeholders with Excel row values.
- Support reusable placeholder naming conventions.

### v0.5 - Certificate Generation Workflow

- Generate certificates from Word templates.
- Save generated Word documents.
- Add batch processing flow from Excel rows.

### v0.6 - PDF Export and Output Packaging

- Export generated certificates as PDF.
- Standardize output naming.
- Prepare release-ready output folders.

### v1.0 - Production Release

- Complete testing.
- Add user documentation.
- Package stable release artifacts.

## Version History

| Version | Date | Summary |
| --- | --- | --- |
| v0.1 | 2026-06-29 | Created project structure, documentation, and VBA module placeholders. |
