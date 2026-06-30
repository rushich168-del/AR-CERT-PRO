# Sample Template Format

AR-CERT PRO expects a Microsoft Word template at the configured default path:

```text
Templates/Certificate_Template.docx
```

The template is opened as a new generated document copy. The original template file is not modified directly.

## Required Placeholders

Placeholders must appear exactly like this in the Word document body:

```text
<<Student Name>>
<<Course>>
<<Certificate No>>
<<Date>>
```

## Example Certificate Text

```text
This certificate is proudly presented to
<<Student Name>>

for successfully completing
<<Course>>

Certificate No: <<Certificate No>>
Date: <<Date>>
```

## Supported Placeholder Rules

- Placeholder format is `<<HeaderName>>`.
- `HeaderName` must match an Excel header from row 1.
- Header names with spaces are supported, such as `<<Student Name>>`.
- Placeholder matching is case-insensitive through Word Find behavior.
- Replacement currently applies to the Word document body.
- Header/footer, shape, text box, PDF-only, and mail merge placeholders are not part of the manual test scope.
