# Changelog

All notable changes to AR-CERT PRO will be documented in this file.

## [v0.2] - 2026-06-30

### Added

- Added v0.2 project configuration constants.
- Added relative path functions based on `ThisDocument.Path`.
- Added timestamped logging with `InitLog`, `WriteLog`, and `WriteErrorLog`.
- Added shared utility helpers for folders, files, filenames, and timestamps.

### Changed

- Updated file/folder helper usage to avoid hardcoded user-specific absolute paths.
- Updated README roadmap status for v0.2.

### Notes

- Excel reading, placeholder replacement, certificate generation, and PDF export workflow are not implemented in this version.

## [v0.1] - 2026-06-29

### Added

- Created complete project folder structure.
- Added initial README documentation.
- Added commercial license notice.
- Added source module placeholders:
  - Main.bas
  - Config.bas
  - ExcelReader.bas
  - PlaceholderEngine.bas
  - PDFExporter.bas
  - Utilities.bas
  - Logger.bas

### Notes

- No business logic has been implemented in this version.
