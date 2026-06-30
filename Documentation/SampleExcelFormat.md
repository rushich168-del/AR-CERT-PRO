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
- Keep one certificate recipient per row.

## Tested Headers

Use this tested sample set for the first real workflow validation:

| Name | Class | Prize |
| --- | --- | --- |
| Aarya Rushi | Class 10 | First Prize |
| Meera Sharma | Class 9 | Excellence Award |
| Arjun Patel | Class 8 | Participation Award |

## Notes

- `Name` is used for output file naming when available.
- If no name-style field is found, AR-CERT PRO uses the Excel row number.
- Blank rows are skipped during batch generation.
