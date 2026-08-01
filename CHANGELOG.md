# Changelog

All notable changes to this project are documented here.
Format loosely follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [Unreleased]

### Added
- **Digital signing script** (`scripts/sign-xlam.ps1`) - creates a self-signed
  "Sandeep Khadka" code-signing certificate and signs the VBA project so the
  Publisher column in Excel's Add-ins dialog shows the publisher's name
  instead of being blank. See docs/installation.md, Option D.

### Fixed
- **Stale VBA signature after signing** - Excel silently invalidates the
  VBA signature whenever it re-saves the add-in after signing (the
  "Remove personal information from file properties on save" privacy
  option is a known trigger), which is why the Publisher column could
  still come up blank. `sign-xlam.ps1` now verifies the signature
  cryptographically (digest inside `xl/vbaProjectSignature.bin` must
  match the current `xl/vbaProject.bin` bytes) and locks the signed
  add-in file read-only so Excel cannot re-save and break it again
  (`-NoLock` opt-out).
- **Add-in registered read-only** - the installer wrote the Excel `OPEN`
  registry value with a `/R` flag, so Excel always loaded the add-in in
  read-only mode. That silently blocked saving after signing the VBA
  project ("file is read-only"), which is why the Publisher name never
  stuck. The installer now registers the add-in without `/R`, and any
  existing `/R` flag is stripped.
- **Signing script verification** - the script now verifies the signature
  from the saved file itself (`xl/vbaProjectSignature*.bin` inside the
  .xlam) instead of trusting Excel's in-memory `VBProject.Signed` property,
  which is a known false-negative source. Also added a file-lock pre-flight
  check (refuses to run while Excel holds the add-in open, preventing the
  `RPC_E_DISCONNECTED` crash on save) and a crash-forgiving save that still
  verifies the file if the Excel COM instance dies mid-save.

## [1.4.0] - 2026-07-29

### Added
- **Configurable hotkeys** - toggle, history back, and history forward
  hotkeys can now be customised via the registry. Defaults remain
  Ctrl+Shift+H, Ctrl+Shift+Z, and Ctrl+Shift+X.
- **Per-mode colour profiles** - in Crosshair mode, row and column
  highlights can now use different colours. Enabled via a toggle in the
  Appearance group with separate colour galleries for each axis.
- **Ribbon cache clearing on install** - the installer now deletes
  Excel's `.officeUI` cache file, which was the most common cause of the
  ribbon tab not appearing after installation.

### Fixed
- **Ribbon tab not appearing in Excel** - the build script now extracts
  and re-zips the OPC package instead of modifying it in-place, which
  eliminates stale central directory entries that could prevent Excel
  from reading the customUI XML.
- **Installer not clearing ribbon cache** - the installer now proactively
  deletes `Excel.officeUI` and prompts the user to close Excel before
  installation.

### Changed
- `HOTKEY_TOGGLE`, `HOTKEY_HISTORY_BACK`, `HOTKEY_HISTORY_FWD` constants
  replaced with `DEFAULT_HOTKEY_*` constants; actual values now read from
  `Settings.HotkeyToggle` etc. so they can be customised.
- `HighlightEngine.CurrentSignature` now includes per-mode colours and
  style so changing them triggers a CF rebuild.
- `Settings.EffectiveRGB` refactored to use shared `ApplyDarkModeTint`.
- About dialog now shows the actual configured hotkeys.
- Version bumped to 1.4.0.

## [1.3.0] - 2026-07-28

### Added
- **Merged-cell support** - range start/end bounds (`_XLCH_RowEnd`, `_XLCH_ColEnd`) ensure highlights cover the full height and width of merged cell ranges.
- **Dynamic GDI colour swatches** - `OleCreatePictureIndirect` and GDI API bitmap generator renders solid colour icons for recent custom colours in the Ribbon gallery.
- **Dark mode support** - toggleable Dark Mode tinting for comfortable contrast on dark Office themes.
- **64-bit Office / VBA7 API compatibility** - updated Win32 API declarations (`ChooseColorAPI`, `FindWindow`, GDI functions) with `PtrSafe` and `LongPtr`.
- **Improved build script diagnostics** - explicit validation and clear error messaging in `build-xlam.ps1`, `export-vba.ps1`, and `import-vba.ps1` when VBA project object model access is blocked.

### Fixed
- Fixed blank icon swatches in recent colours Ribbon gallery.
- Fixed single-cell constraint on merged cell selections.

## [1.2.0] - 2026-07-20

### Added
- **Workbook exclusion list** - per-workbook toggle via `_XLCH_Excluded`.
- **Worksheet exclusion list** - per-sheet toggle via `_XLCH_SheetExcluded`.
- **Keyboard shortcuts** - `Ctrl+Shift+H` (toggle), `Ctrl+Shift+Z/X` (history).
- **Status bar integration** - shows mode, colour, and extras.
- **Recent colours in gallery** - last 4 custom colours in the gallery.
- **Native colour picker** - Windows API `ChooseColor`, no palette borrowing.
- **Border-only highlight style** - thick borders instead of fill.
- **Intersection accent** - stronger tint at row/column crossing.
- **Protected-sheet support (opt-in)** - temporarily unprotects to apply CF.
- **Animated highlight pulse** - 3-step pulse on selection change.
- **Selection history** - back/forward navigation of last 20 cells.
- **Profiles** - named bundles of settings with a dropdown and save button.

### Changed
- Colour gallery is now dynamic to support recent colours.
- `HighlightEngine` now supports fill vs border, intersection, animation, protected sheets.
- Build scripts import `ColourPicker.bas`, `SelectionHistory.bas`, `Profiles.bas`.

### Known issues
- Highlight range is bounded to UsedRange ∪ VisibleRange.
- Recent colour gallery items show blank swatches (needs `OleCreatePictureIndirect`).

## [1.1.0] - Unreleased

### Added
- **Workbook exclusion list** - per-workbook toggle via a hidden defined name
  (`_XLCH_Excluded`) that travels with the file. Accessible from the Options
  group on the ribbon. Excluded workbooks are skipped entirely by the
  highlighter.
- **Worksheet exclusion list** - per-sheet toggle via `_XLCH_SheetExcluded`
  worksheet-scoped defined name. Accessible from the Options group. Works
  independently of the workbook-level exclusion.
- **Keyboard shortcut** - `Ctrl+Shift+H` toggles the highlighter on/off
  globally. Registered via `Application.OnKey` alongside the event sink
  lifecycle in `AddinHost`.
- **Status bar integration** - shows current mode and colour (e.g.
  "XLCH: Crosshair - Yellow") in the status bar as a quick visual reference.
- **Recent colours in gallery** - the last 4 custom RGB values are remembered,
  persisted in the registry, and surfaced as additional items in the colour
  gallery when the user picks a custom colour.
- **Native colour picker** - replaced the old `xlDialogEditColor` approach
  (which borrowed palette slot 56 from the active workbook) with the Windows
  Common Dialog API (`ChooseColor` from `comdlg32.dll`). No workbook palette
  slots are touched or modified.

### Changed
- `EnsureNamesExist`, `RebuildConditionalFormatting`,
  `RemoveOurConditionalFormatting`, and `UntrackSheet` in `HighlightEngine`
  changed from `Private` to `Public` so the exclusion callbacks can call them.
- Colour gallery is now dynamic (`getItemCount`/`getItemID`/`getItemImage`)
  to support recent colours as additional items.
- `AddinHost.ShutDown` now unregisters the hotkey before tearing down the
  event sink.
- `PrompForCustomRGB` in `RibbonCallbacks` replaced with call to
  `ColourPicker.PromptForCustomRGB` (Windows API).
- About dialog now mentions the hotkey.
- Build scripts import the new `ColourPicker.bas` module.

### Removed
- Old `xlDialogEditColor`-based colour picker (`PromptForCustomRGB` in
  `RibbonCallbacks`) - superseded by `ColourPicker.bas`.

### Known issues
- Highlight range is bounded to UsedRange ∪ VisibleRange rather than the
  full grid - see docs/architecture.md for why, and the resulting edge case
  around scrolling into distant empty space.
- Recent colour gallery items show blank swatches (the `GenerateColourSwatch`
  helper currently returns Nothing - an `OleCreatePictureIndirect`-based
  implementation would render solid-colour bitmaps).

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
