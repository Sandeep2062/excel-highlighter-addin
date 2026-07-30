# Future features / roadmap

Rough ideas, not commitments, roughly ordered by how likely they are to
actually be worth doing next.

## Implemented in v1.4.0

- **Configurable hotkeys** - the toggle, history back, and history forward
  hotkeys can now be customised via the registry (SaveSetting). Defaults
  remain Ctrl+Shift+H, Ctrl+Shift+Z, and Ctrl+Shift+X.
- **Per-mode colour profiles** - in Crosshair mode, the row and column
  highlights can now use different colours. Enabled via a toggle in the
  Appearance group, with separate colour galleries for each axis.
- **Ribbon cache clearing** - the installer now deletes Excel's `.officeUI`
  cache file, which was the most common cause of the ribbon tab not
  appearing after installation.
- **Build script reliability** - the build script now extracts and re-zips
  the OPC package instead of modifying it in-place, eliminating stale
  central directory entries that could prevent Excel from reading the
  customUI XML.

## Implemented in v1.3.0

- **Merged-cell support** - active selection detection uses range bounds (`_XLCH_Row`, `_XLCH_RowEnd`, `_XLCH_Col`, `_XLCH_ColEnd`), ensuring crosshairs cover full merged cell dimensions.
- **Dynamic GDI swatches** - recent custom colours render solid 32x32 swatches in the Ribbon gallery via `OleCreatePictureIndirect` and GDI bitmap generation.
- **Dark mode support** - toggleable Dark Mode tinting for comfortable visual contrast on dark Office themes.
- **64-bit Office compatibility** - complete `PtrSafe` / `LongPtr` API coverage for Windows Common Dialog and GDI functions.

## Implemented in v1.2.0

- **Workbook exclusion list** - a per-workbook "don't highlight this one" toggle via `_XLCH_Excluded`.
- **Worksheet exclusion list** - per-sheet toggle via `_XLCH_SheetExcluded`.
- **Keyboard shortcuts** - `Ctrl+Shift+H` (toggle), `Ctrl+Shift+Z/X` (history).
- **Status bar integration** - shows mode, colour, and active extras.
- **Recent colours in gallery** - last 4 custom colours in the gallery.
- **Native colour picker** - Windows API `ChooseColor`, no palette borrowing.
- **Border-only highlight style** - thick borders instead of fill.
- **Intersection accent** - stronger tint at row/column crossing.
- **Protected-sheet support (opt-in)** - temporarily unprotects to apply CF.
- **Animated highlight pulse** - 3-step pulse on selection change.
- **Selection history** - back/forward navigation of last 20 cells.
- **Profiles** - named bundles of settings with a dropdown and save button.

## Fairly likely (next)

- **Hotkey configuration UI** - currently hotkeys are configurable via the
  registry only; a ribbon dialog or InputBox flow would make this more
  user-friendly.
- **Per-mode custom colours** - the per-mode colour galleries currently
  only offer the preset palette. Allowing custom RGB for each axis would
  complete the feature.

## Explicitly not planned

- Anything that would require permanently modifying `Interior.Color` or other real cell formatting. That's a hard architectural boundary for this project, not just a current limitation.
