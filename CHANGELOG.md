# Changelog

All notable changes to this project are documented here.
Format loosely follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [1.0.0] - Unreleased

### Added
- Initial release: Row / Column / Crosshair highlighting via non-destructive
  conditional formatting, driven by Application-level events.
- Ribbon tab ("Highlighter") with enable/disable toggle, mode buttons, a
  seven-colour gallery, custom RGB colour picker, reset-to-defaults, and an
  About dialog.
- Persisted preferences (enabled state, mode, colour) via SaveSetting/
  GetSetting, scoped to the current Windows user.
- File-based error/info logging under `%APPDATA%\ExcelCrosshairHighlighter`.
- Automatic cleanup of all added formatting/names on workbook close and on
  add-in uninstall.
- PowerShell scripts for exporting/importing the VBA project and building
  the `.xlam` from source.

### Known issues
- Highlight range is bounded to UsedRange ∪ VisibleRange rather than the
  full grid - see docs/architecture.md for why, and the resulting edge case
  around scrolling into distant empty space.
- Custom colour picker temporarily borrows workbook colour palette slot 56.
