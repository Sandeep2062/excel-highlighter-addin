# Excel-Highlighter

A small Excel add-in that highlights the active row, column, or both (a
"crosshair"), following the active cell as you move around a workbook. It
runs globally once installed - no per-workbook setup, no macros to enable in
individual files.

It does **not** touch `Interior.Color` or any other cell formatting. The
highlight is a temporary conditional-formatting overlay that is added and
removed as needed, and is stripped out entirely when a workbook closes or the
add-in is unloaded. If you save a workbook while the highlighter is running,
nothing highlighter-related ends up saved in the file's own formatting.

## Features

- Four modes: **Row**, **Column**, **Crosshair** (row + column together), **Cell** (active cell only), or off.
- Seven preset colours plus a native Windows RGB picker (64-bit and 32-bit compatible).
- **Per-mode colour profiles** - in Crosshair mode, row and column highlights can use different colours.
- **Configurable hotkeys** - toggle, history back, and history forward keys can be customised via the registry.
- Full merged-cell support: highlights automatically cover the entire dimensions of merged ranges.
- Dark mode tinting for pleasant visual contrast on dark Office themes.
- Dynamic runtime GDI swatches for recent custom colours in the Ribbon gallery.
- Workbook and sheet exclusion toggles stored in file-level defined names.
- Selection history navigation (default `Ctrl+Shift+Z` back, `Ctrl+Shift+X` forward).
- Named profiles for quick configuration switching.
- Works across every open workbook and window at once.
- Settings persist between Excel sessions.
- No VBA project password, no hidden macros required in your own workbooks.
- Designed to degrade gracefully on protected sheets and chart sheets rather than throwing errors at you.

## How it works, briefly

Four hidden, workbook-scoped defined names (`_XLCH_Row`, `_XLCH_RowEnd`, `_XLCH_Col`, `_XLCH_ColEnd`) hold the active selection bounds as plain numbers. Conditional formatting rules on each worksheet reference those names (`=AND(ROW()>=_XLCH_Row,ROW()<ShortEnd)`, etc.). Moving the active cell or selecting merged ranges updates the four names - a cheap operation - rather than adding or removing formatting rules on every keystroke. The CF rules themselves are only (re)built the first time a sheet is visited, or when you change mode, colour, style, or options from the ribbon.

See [docs/architecture.md](docs/architecture.md) for the full picture, including why the highlight range is bounded rather than applied to entire 1,048,576-row sheets.

## User Guide & Documentation

For a comprehensive step-by-step feature reference, shortcuts, and usage examples, see [docs/user-guide.md](docs/user-guide.md).

Quick reference:
- **Toggle Highlight**: Click **Highlight: On / Off** on the Ribbon or press `Ctrl+Shift+H` (configurable).
- **Mode Selection**: Click **Row**, **Column**, **Crosshair**, or **Cell**.
- **Per-Mode Colours**: Toggle **Per-Mode Colours** in the Appearance group, then pick separate colours for row and column.
- **Navigate Cell History**: Press `Ctrl+Shift+Z` (back) or `Ctrl+Shift+X` (forward) - both configurable.
- **Dark Mode**: Click **Dark Mode** under Options for comfortable dark-theme reading.

## Installation

- **Option 1 (Recommended 1-Click)**: Double-click `install.bat` (or run `install.ps1` in PowerShell). It automatically builds `excel-highlighter.xlam`, copies the add-in and `images/` icons to `%APPDATA%\Microsoft\AddIns\ExcelHighlighter`, activates it in the Windows Registry for Excel, and clears Excel's ribbon cache so the Highlighter tab appears immediately.
- **Option 2 (Manual)**: See [docs/installation.md](docs/installation.md) for step-by-step instructions.

**Ribbon tab not appearing?** The installer now clears Excel's `.officeUI` cache automatically. If you still don't see the Highlighter tab after installation: close Excel completely, re-run `install.bat`, then open Excel again.

To uninstall anytime, double-click `uninstall.bat` (or run `uninstall.ps1`).

## Folder structure

```
excel-highlighter-addin/
├── src/                      VBA source, exported as text (.cls / .bas)
│   ├── ThisWorkbook.cls       add-in workbook lifecycle
│   ├── AddinHost.bas          owns the EventApp instance's lifetime
│   ├── EventApp.cls           Application-level WithEvents sink
│   ├── HighlightEngine.bas    conditional-formatting overlay logic
│   ├── RibbonCallbacks.bas    every customUI14.xml callback
│   ├── Settings.bas           persisted preferences (SaveSetting/GetSetting)
│   ├── ColourPicker.bas       Windows Common Dialog API & GDI swatch generator
│   ├── SelectionHistory.bas    cell navigation history stack
│   ├── Profiles.bas           named settings profiles
│   ├── Utilities.bas          small stateless helpers
│   ├── Logging.bas            file-based error/info logging
│   └── Constants.bas          shared literals and enums
├── customUI/
│   ├── customUI14.xml         RibbonX definition (Office 2010+ / customUI14)
│   └── images/                colour swatch PNGs loaded by the ribbon
├── docs/
│   ├── user-guide.md          complete feature & usage reference guide
│   ├── installation.md        installation & troubleshooting guide
│   ├── architecture.md        technical architecture & performance notes
│   └── future-features.md     project roadmap & feature backlog
├── scripts/                   PowerShell helpers for exporting/importing
│                               the VBA project and building the .xlam
├── tests/
│   └── manual-test-checklist.md
├── CHANGELOG.md
├── LICENSE
└── README.md
```

## Development workflow

The VBA project itself lives inside `excel-highlighter.xlam`, which is a binary Office file and isn't diff-friendly in git. The `.cls`/`.bas` files under `src/` are the text-exported source of truth:

1. Open the `.xlam` in Excel, make your changes in the VBA editor (Alt+F11).
2. Run `scripts/export-vba.ps1` to write the current VBA project back out to `src/` as text.
3. Commit the exported text files.

Going the other direction (text → binary) after a fresh checkout:

1. Run `scripts/build-xlam.ps1`, which creates a blank `.xlam`, imports every module from `src/`, sets the RibbonX customization from `customUI/customUI14.xml`, and saves it.

Both scripts drive Excel via its COM object model (`CreateObject("Excel.Application")`) and the VBA Extensibility library, so Excel needs to be installed on the machine running them, and "Trust access to the VBA project object model" needs to be enabled under Excel Options → Trust Center → Macro Settings.

## Known limitations

- The highlight is bounded to each sheet's used range plus whatever is currently visible on screen, not the entire 1,048,576 × 16,384 grid. On a sheet with very sparse, widely separated data, scrolling into a completely empty area far from both isn't highlighted until you interact with it. See `docs/architecture.md` for the reasoning.
- Conditional formatting rules count against Excel's per-sheet CF limit. This add-in only ever adds one to three rules per sheet, but if a workbook already has an unusually large number of existing rules from other sources, adding ours could push it toward that ceiling.
- Sheets protected without "Allow formatting cells" enabled won't be highlighted unless "Allow Protected" mode is enabled in the Options group.

## Roadmap

See [docs/future-features.md](docs/future-features.md).

## FAQ

**Does this modify my workbook file?**
No. The conditional formatting and defined names are added at runtime and
removed again on `WorkbookBeforeClose` / add-in uninstall. If you never close
the workbook through Excel (e.g. the process is killed), whatever CF/names
were live at that point could end up saved - this is the one edge case worth
knowing about.

**Will it slow down a huge workbook?**
The intent is no, or not noticeably - see the architecture notes on why the
highlight range is bounded. If you do notice a slowdown on a specific
workbook, please open an issue with a rough description of its size/shape.

**Can I use a different colour per mode?**
Yes - enable **Per-Mode Colours** in the Appearance group on the ribbon.
This gives you separate colour galleries for the row and column highlights
in Crosshair mode.

## Contributing

Issues and pull requests are welcome. If you're proposing a new ribbon
control or a new highlight mode, please open an issue first to talk through
the approach - the CF-overlay design has some non-obvious constraints (see
architecture.md) that are worth checking against before writing code.

## License

MIT - see [LICENSE](LICENSE).
