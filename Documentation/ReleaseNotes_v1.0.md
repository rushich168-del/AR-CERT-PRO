# AR-CERT PRO v1.0 Release Notes

## Features

- Microsoft Word VBA certificate generation workflow.
- Excel-driven recipient data source using late binding.
- Word template copy creation without modifying the original template.
- Placeholder replacement using `<<HeaderName>>` syntax.
- DOCX output generation.
- PDF export using Word fixed-format export.
- Batch processing for Excel data rows.
- Output filename sanitization.
- Timestamped logs under `Output/Logs`.
- Manual test runner for workflow validation.

## Tested Result

v1.0 is based on the real Microsoft Word VBA validation completed in v0.9.

- TestLogger: PASS
- TestConfigPaths: PASS
- TestExcelEngine: PASS
- TestWordEngine: PASS
- TestPDFExporter: PASS
- TestCompleteWorkflow: PASS
- Total rows: 3
- Success count: 3
- Failure count: 0
- Generated DOCX files successfully.
- Generated PDF files successfully.

## Known Limitations

- No GUI is included.
- No mail merge feature is included.
- Placeholder replacement is focused on Word document body content.
- The default worksheet name is `Sheet1`.
- The default Excel path is `Excel/Students.xlsx`.
- The default template path is `Templates/Certificate_Template.docx`.

## Next Roadmap

- Professional user interface.
- Progress bar.
- Settings window for paths and worksheet name.
- Installer or packaged distribution.
- Branding and polished user documentation.
