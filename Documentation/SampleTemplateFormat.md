# Sample Template Format

AR-CERT PRO expects a Microsoft Word template at the configured default path:

```text
Templates/Certificate_Template.docx
```

The template is opened as a new generated document copy. The original template file is not modified directly.

## Required Placeholders

For the tested sample workbook, placeholders must appear exactly like this in the Word document body:

```text
<<Name>>
<<Class>>
<<Prize>>
```

## Example Certificate Text

```text
This certificate is proudly presented to
<<Name>>

from
<<Class>>

for receiving
<<Prize>>
```

## Supported Placeholder Rules

- Placeholder format is `<<HeaderName>>`.
- `HeaderName` must match an Excel header from row 1.
- Header names with spaces are supported by the engine, but the tested sample uses `Name`, `Class`, and `Prize`.
- Placeholder matching is case-insensitive through Word Find behavior.
- Replacement currently applies to the Word document body.
- Header/footer, shape, text box, PDF-only, and mail merge placeholders are not part of the manual test scope.
