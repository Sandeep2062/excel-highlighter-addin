# Excel Crosshair Highlighter

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

- Three modes: **Row**, **Column**, **Crosshair** (row + column together), or off.
- Seven preset colours plus a custom RGB picker.
- Works across every open workbook and window at once.
- Settings (enabled state, mode, colour) persist between Excel sessions.
- No VBA project password, no hidden macros required in your own workbooks.
- Designed to degrade gracefully on protected sheets and chart sheets rather
  than throwing errors at you.

## How it works, briefly

Two hidden, workbook-scoped defined names (`_XLCH_Row`, `_XLCH_Col`) hold the
active row and column as plain numbers. Conditional formatting rules on each
worksheet reference those names (`=ROW()=_XLCH_Row`, etc.). Moving the active
cell just updates the two names - a cheap operation - rather than adding or
removing formatting rules on every keystroke. The CF rules themselves are
only (re)built the first time a sheet is visited, or when you change mode,
colour, or the enabled toggle from the ribbon.

See [docs/architecture.md](docs/architecture.md) for the full picture,
including why the highlight range is bounded rather than applied to entire
1,048,576-row sheets.

## Installation

See [docs/installation.md](docs/installation.md). Short version: build or
download `excel-crosshair-highlighter.xlam`, keep the `images` folder next to
it, then add it via Excel Options → Add-ins → Manage Excel Add-ins → Browse.

## Folder structure

```
excel-crosshair-highlighter/
├── src/                      VBA source, exported as text (.cls / .bas)
│   ├── ThisWorkbook.cls       add-in workbook lifecycle
│   ├── AddinHost.bas          owns the EventApp instance's lifetime
│   ├── EventApp.cls           Application-level WithEvents sink
│   ├── HighlightEngine.bas    conditional-formatting overlay logic
│   ├── RibbonCallbacks.bas    every customUI14.xml callback
│   ├── Settings.bas           persisted preferences (SaveSetting/GetSetting)
│   ├── Utilities.bas          small stateless helpers
│   ├── Logging.bas            file-based error/info logging
│   └── Constants.bas          shared literals and enums
├── customUI/
│   ├── customUI14.xml         RibbonX definition (Office 2010+ / customUI14)
│   └── images/                colour swatch PNGs loaded by the ribbon
├── docs/
│   ├── installation.md
│   ├── architecture.md
│   └── future-features.md
├── scripts/                   PowerShell helpers for exporting/importing
│                               the VBA project and building the .xlam
├── tests/
│   └── manual-test-checklist.md
├── CHANGELOG.md
├── LICENSE
└── README.md
```

## Development workflow

The VBA project itself lives inside `excel-crosshair-highlighter.xlam`, which
is a binary Office file and isn't diff-friendly in git. The `.cls`/`.bas`
files under `src/` are the text-exported source of truth:

1. Open the `.xlam` in Excel, make your changes in the VBA editor (Alt+F11).
2. Run `scripts/export-vba.ps1` to write the current VBA project back out to
   `src/` as text.
3. Commit the exported text files.

Going the other direction (text → binary) after a fresh checkout:

1. Run `scripts/build-xlam.ps1`, which creates a blank `.xlam`, imports every
   module from `src/`, sets the RibbonX customization from
   `customUI/customUI14.xml`, and saves it.

Both scripts drive Excel via its COM object model (`CreateObject("Excel.Application")`)
and the VBA Extensibility library, so Excel needs to be installed on the
machine running them, and "Trust access to the VBA project object model"
needs to be enabled under Excel Options → Trust Center → Macro Settings.

## Known limitations

- The highlight is bounded to each sheet's used range plus whatever is
  currently visible on screen, not the entire 1,048,576 × 16,384 grid. On a
  sheet with very sparse, widely separated data, scrolling into a completely
  empty area far from both isn't highlighted until you interact with it. See
  `docs/architecture.md` for the reasoning.
- Conditional formatting rules count against Excel's per-sheet CF limit. This
  add-in only ever adds one or two rules per sheet, but if a workbook already
  has an unusually large number of existing rules from other sources, adding
  ours could push it toward that ceiling.
- Sheets protected without "Allow formatting cells" enabled won't be
  highlighted; there's no reliable way around that without actually
  unprotecting the sheet, which this add-in intentionally never does on your
  behalf.
- The custom colour picker uses Excel's own `xlDialogEditColor` dialog via a
  scratch palette slot. On very old custom-palette-heavy workbooks this is
  usually fine but is worth knowing about if slot 56 is meaningfully used
  elsewhere in that specific workbook.

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
Not currently - colour is global, mode is global. Per-mode colour is a
reasonable idea; see the roadmap.

## Contributing

Issues and pull requests are welcome. If you're proposing a new ribbon
control or a new highlight mode, please open an issue first to talk through
the approach - the CF-overlay design has some non-obvious constraints (see
architecture.md) that are worth checking against before writing code.

## License

MIT - see [LICENSE](LICENSE).
