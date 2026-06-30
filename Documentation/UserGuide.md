# AR-CERT PRO User Guide

## Setup

1. Clone or download the AR-CERT PRO repository.
2. Open Microsoft Word.
3. Create or open a macro-enabled Word document (`.docm`) to act as the runner.
4. Press `Alt+F11` to open the VBA editor.
5. Import all modules from the `Source` folder.
6. Save the runner document.
7. Enable macros when prompted by Microsoft Word.

## Required Excel Format

Create the Excel workbook at:

```text
Excel/Students.xlsx
```

Use worksheet:

```text
Sheet1
```

The tested sample uses these headers in row 1:

| Name | Class | Prize |
| --- | --- | --- |
| Aarya Rushi | Class 10 | First Prize |
| Meera Sharma | Class 9 | Excellence Award |
| Arjun Patel | Class 8 | Participation Award |

Data must begin on row 2. Blank rows are skipped.

## Required Word Template Format

Create the Word template at:

```text
Templates/Certificate_Template.docx
```

Use placeholders in the document body that match Excel headers exactly:

```text
<<Name>>
<<Class>>
<<Prize>>
```

The original template is opened as a copy during generation and is not modified directly.

## Running Macro

Run the main workflow from the Word VBA editor:

```text
GenerateCertificates
```

For validation, run:

```text
TestCompleteWorkflow
```

## Output Files

Generated Word documents are saved in:

```text
Output/
```

Generated PDF files are saved in:

```text
Output/
```

Output filenames use the `Name` field when available. If no name field is found, the Excel row number is used.

## Log Files

Log files are saved in:

```text
Output/Logs/
```

Logs include test results, row-level success or failure, and batch summary counts.

## Troubleshooting

### Macros do not run

Confirm the runner document is saved as `.docm` and macros are enabled in Microsoft Word.

### Excel workbook is not found

Confirm the workbook exists at `Excel/Students.xlsx` relative to the project folder.

### Worksheet is not found

Confirm the worksheet is named `Sheet1`.

### Placeholders are not replaced

Confirm placeholders are in the Word document body and match Excel headers exactly, including spaces and spelling.

### PDF is not generated

Confirm Microsoft Word can save/export documents and that the `Output` folder is writable.

### Review failures

Open the newest log file in `Output/Logs/` and look for `FAIL` or `ERROR` entries.
