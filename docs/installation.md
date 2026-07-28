# Installation

## Option A - build it yourself (recommended while this is source-only)

Requirements: Excel 2024 (or any recent Excel 2010+ build - RibbonX
`customUI14` and everything else used here has been available since Office
2010) on Windows, with "Trust access to the VBA project object model"
enabled (File → Options → Trust Center → Trust Center Settings → Macro
Settings).

1. Clone the repository.
2. Run `scripts/build-xlam.ps1` from PowerShell. This will:
   - launch Excel via COM,
   - create a new blank workbook,
   - import every module under `src/`,
   - apply the RibbonX customization from `customUI/customUI14.xml`,
   - save it as `excel-crosshair-highlighter.xlam` in the repository root.
3. Copy `excel-crosshair-highlighter.xlam` **and** the `customUI/images`
   folder to wherever you want the add-in to live permanently, e.g.
   `%APPDATA%\Microsoft\AddIns\ExcelCrosshairHighlighter\`. Rename the copied
   `images` folder to sit alongside the `.xlam` (i.e.
   `...\ExcelCrosshairHighlighter\images\*.png`) - the ribbon loads the
   colour swatches relative to the add-in's own file location at runtime.
4. In Excel: File → Options → Add-ins → at the bottom, set "Manage" to
   "Excel Add-ins" → Go... → Browse... → select the `.xlam` you copied in
   step 3 → OK, and make sure its checkbox is ticked.
5. You should see a new **Highlighter** tab on the ribbon.

## Option B - manual VBA import (no PowerShell)

1. Create a new blank workbook.
2. Alt+F11 to open the VBA editor.
3. File → Import File... for each `.bas` file in `src/`
   (`Constants.bas`, `Logging.bas`, `Utilities.bas`, `Settings.bas`,
   `HighlightEngine.bas`, `AddinHost.bas`, `RibbonCallbacks.bas`).
4. For `EventApp.cls`: Insert → Class Module, rename it to `EventApp`, then
   paste in the contents of `src/EventApp.cls` (skip the `VERSION 1.0 CLASS`
   header block - that's metadata Excel generates itself).
5. Double-click `ThisWorkbook` in the Project Explorer and paste in the
   contents of `src/ThisWorkbook.cls` (again, skip the header block).
6. Attach the RibbonX customization. The VBA editor can't do this directly -
   use a tool such as the free [Custom UI Editor for Microsoft Office
   (RibbonX Editor)](https://github.com/fernandreu/office-ribbonx-editor):
   open your workbook in it, add `customUI/customUI14.xml` as the "Office
   2010+" custom UI part, and for each `<item>`/`<button>` that uses
   `getImage`, embed the matching PNG from `customUI/images/` as an image
   resource with the same id you see referenced (or just leave the runtime
   `LoadPicture`-based loading in `RibbonCallbacks.GetSwatchImage` as-is,
   which reads the PNGs from disk instead - that's what this add-in does by
   default).
7. File → Save As → choose "Excel Add-in (*.xlam)" as the file type.
8. Follow steps 3-5 from Option A to install it.

## Uninstalling / disabling

Untick the add-in in File → Options → Add-ins → Manage Excel Add-ins. This
fires `Workbook_AddinUninstall`, which removes every conditional formatting
rule and defined name the add-in added to any currently open workbook before
it unloads.

## Troubleshooting

- **No "Highlighter" tab appears** - check that the `.xlam` is ticked (not
  just listed) under Add-ins, and that macros/add-ins aren't blocked by
  Group Policy or the Trust Center's "Disable all macros" setting.
- **Colour swatches show as blank squares** - the `images` folder isn't
  sitting next to the `.xlam`, or was renamed. See step 3 above.
- **Nothing highlights on a specific sheet** - it's very likely protected
  without "Allow formatting cells" turned on; see the Known Limitations
  section of the main README.
- **Errors seem to be happening silently** - check
  `%APPDATA%\ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log`, which
  the add-in writes to instead of showing message boxes for unexpected
  errors.
