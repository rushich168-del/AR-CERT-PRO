# AR-CERT PRO

## Project Description

AR-CERT PRO is a Microsoft Word VBA application designed to generate professional certificates automatically from Excel data. The project is structured as a commercial-quality automation solution for AaryaRushi Automation Labs, with a clean module layout, reusable components, and a roadmap for controlled feature development.

Current Version: **v0.9**

## Features

- Word-based certificate generation workflow.
- Excel-driven recipient data source.
- Late-bound Excel reader engine.
- Word template engine that opens templates as generated document copies.
- Placeholder replacement for `<<HeaderName>>` fields in document body content.
- PDF export engine using Word fixed-format export.
- Batch certificate generation from Excel rows.
- Template-based certificate design support.
- Placeholder replacement architecture.
- PDF export module boundary.
- Configuration constants and relative project path helpers.
- Timestamped logging under `Output/Logs`.
- Shared file, folder, filename, and timestamp utilities.
- Organized project structure ready for source control and release packaging.

> Note: v0.9 validates the real Microsoft Word VBA workflow. Mail merge and GUI work remain planned for future versions.

## Folder Structure

```text
AR-CERT-PRO/
|-- Documentation/
|   |-- SampleExcelFormat.md
|   `-- SampleTemplateFormat.md
|-- Source/
|   |-- Main.bas
|   |-- Config.bas
|   |-- ExcelReader.bas
|   |-- PlaceholderEngine.bas
|   |-- PDFExporter.bas
|   |-- TestRunner.bas
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

### v0.3 - Excel Reader Engine - Complete

- Initialize and close Excel using late binding.
- Open configured workbooks safely.
- Select worksheets and read cells/ranges.

### v0.4 - Word Template Engine - Complete

- Open configured Word templates as new generated document copies.
- Save generated Word documents safely.
- Close and expose the active generated document.

### v0.5 - Placeholder Replacement Engine - Complete

- Replace `<<HeaderName>>` placeholders in Word document body content.
- Support Excel headers with spaces, such as `<<Student Name>>`.
- Replace all placeholders from header and row-value arrays.

### v0.6 - PDF Export Engine - Complete

- Export generated Word documents to PDF.
- Ensure PDF output folders exist before export.
- Sanitize PDF output filenames.

### v0.7 - Batch Certificate Generator - Complete

- Generate certificates from Word templates.
- Save generated Word documents.
- Add batch processing flow from Excel rows.

### v0.8 - Manual Testing Package - Complete

- Document sample Excel workbook structure.
- Document sample Word template placeholders.
- Add VBA test runner for manual end-to-end checks.

### v0.9 - Real VBA Test Validation - Complete

- Validated the complete workflow inside Microsoft Word VBA.
- Generated DOCX and PDF output files successfully.
- Confirmed 3 data rows processed with 3 successes and 0 failures.

### v1.0 - Output Packaging - Planned

- Organize generated DOCX and PDF outputs.
- Standardize output naming.
- Prepare release-ready output folders.

### v1.0 - Production Release - Planned

- Complete testing.
- Add user documentation.
- Package stable release artifacts.

## Manual Testing Guide

### 1. Import VBA modules

Open the Word VBA editor, create or open the AR-CERT PRO macro-enabled Word document, and import every `.bas` file from the `Source` folder.

### 2. Enable macros

Save the Word file as a macro-enabled document (`.docm`) and enable macros when Word prompts for permission.

### 3. Prepare Excel file

Create the workbook at:

```text
Excel/Students.xlsx
```

Use worksheet `Sheet1`. Row 1 must contain the sample headers from `Documentation/SampleExcelFormat.md`.

### 4. Prepare Word template

Create the template file at:

```text
Templates/Certificate_Template.docx
```

Add placeholders in the document body using the format shown in `Documentation/SampleTemplateFormat.md`, such as `<<Name>>`, `<<Class>>`, and `<<Prize>>`.

### 5. Run the workflow test

In the VBA editor, run:

```text
TestCompleteWorkflow
```

### 6. Expected output

The test writes PASS/FAIL entries to the log and then runs the batch certificate workflow.

Generated DOCX files are saved in:

```text
Output/
```

Generated PDF files are saved in:

```text
Output/
```

Log files are saved in:

```text
Output/Logs/
```

## Real VBA Test Validation

v0.9 was tested successfully inside Microsoft Word VBA.

- `TestLogger`: PASS
- `TestConfigPaths`: PASS
- `TestExcelEngine`: PASS
- `TestWordEngine`: PASS
- `TestPDFExporter`: PASS
- `TestCompleteWorkflow`: PASS
- Batch total rows: 3
- Batch success count: 3
- Batch failure count: 0
- Generated DOCX files successfully.
- Generated PDF files successfully.

## Version History

| Version | Date | Summary |
| --- | --- | --- |
| v0.9 | 2026-06-30 | Validated real Word VBA workflow, fixed WordEngine compile issue, and aligned sample placeholder docs. |
| v0.8 | 2026-06-30 | Added manual testing package with sample formats and VBA test runner. |
| v0.7 | 2026-06-30 | Added batch certificate generator connecting Excel, Word, placeholder, DOCX, and PDF engines. |
| v0.6 | 2026-06-30 | Added PDF export engine using Word fixed-format export. |
| v0.5 | 2026-06-30 | Added placeholder replacement engine for Word document body placeholders. |
| v0.4 | 2026-06-30 | Added Word template engine for opening, saving, closing, and accessing generated documents. |
| v0.3 | 2026-06-30 | Added late-bound Excel reader engine with workbook, worksheet, cell, and range access. |
| v0.2 | 2026-06-30 | Added configuration constants, relative path helpers, logging, and shared utilities. |
| v0.1 | 2026-06-29 | Created project structure, documentation, and VBA module placeholders. |
