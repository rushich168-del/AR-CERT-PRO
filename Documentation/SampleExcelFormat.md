# Sample Excel Format

AR-CERT PRO expects a Microsoft Excel workbook at the configured default path:

```text
Excel/Students.xlsx
```

The default worksheet name is:

```text
Sheet1
```

## Required Structure

- Row 1 must contain headers.
- Data must start on row 2.
- Each header maps directly to a Word placeholder with the same text inside `<<` and `>>`.
- Header names may contain spaces.
- Keep one certificate recipient per row.

## Example Headers

| Student Name | Course | Certificate No | Date |
| --- | --- | --- | --- |
| Aarya Rushi | Word VBA Automation | CERT-001 | 2026-06-30 |
| Meera Sharma | Excel Reporting | CERT-002 | 2026-06-30 |
| Arjun Patel | Office Automation | CERT-003 | 2026-06-30 |
| Kavya Rao | Document Generation | CERT-004 | 2026-06-30 |
| Rohan Verma | Certificate Workflow | CERT-005 | 2026-06-30 |

## Notes

- `Student Name` is used for output file naming when available.
- If no name-style field is found, AR-CERT PRO uses the Excel row number.
- Blank rows are skipped during batch generation.
