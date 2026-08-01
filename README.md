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

## How it works

The add-in creates four hidden, workbook-scoped defined names (`XLCH_Row`, `XLCH_RowEnd`, `XLCH_Col`, `XLCH_ColEnd`) that store the active selection bounds as plain numbers. Temporary non-destructive Conditional Formatting rules are applied to each active worksheet referencing those defined names (`=AND(ROW()>=XLCH_Row,ROW()<=XLCH_RowEnd)`). 

When you move your selection or pick merged cells, the defined names are updated instantaneously—a lightweight, non-blocking operation—without constantly deleting and re-creating formatting rules on every arrow key press or mouse click.

### Key Functions & Integration

- **Toggle Modes**: Switch seamlessly between **Row**, **Column**, **Crosshair** (both row and column), **Cell** (active cell focus), or **Off**.
- **Keyboard Shortcuts**: Press **`Ctrl + Shift + H`** anytime in Excel to toggle highlighting on or off globally.
- **Right-Click Context Menu**: Right-clicking any cell in Excel provides a native **Toggle Highlighter (Ctrl+Shift+H)** context menu item for instant access.
- **Per-Mode Colours**: Enable different colors for rows and columns simultaneously in Crosshair mode.
- **Dark Mode Tinting**: Automatically adjusts color saturation for comfortable reading on dark Office themes.
- **Selection Navigation History**: Jump backward (`Ctrl+Shift+Z`) and forward (`Ctrl+Shift+X`) across recent cell locations.
- **Workbook & Sheet Exclusions**: Exclude specific workbooks or sheets from highlighting via Ribbon toggles.

See [docs/architecture.md](docs/architecture.md) for full architectural details.

## Ongoing Issues & Technical Considerations

Here are known technical behaviors, platform-specific edge cases, and active considerations:

1. **64-Bit Office Win32 GDI Callback Crash (Resolved via `imageMso`)**:
   - *Issue*: On 64-bit Office 2021/2024/365, dynamic Ribbon GDI image generation via `OleCreatePictureIndirect` throws a COM exception (`0x8000FFFF E_UNEXPECTED`), which causes Excel's RibbonX engine to silently abort loading the custom tab.
   - *Mitigation*: The custom UI definition (`customUI14.xml`) uses native Office `imageMso` icons (`ColorPalette`, `DiagramColorToggle`), and GDI callbacks have fail-safe error trapping to ensure the Ribbon tab renders reliably on all 32-bit and 64-bit Office builds.

2. **Defined Name Syntax Rules in Modern Excel 2021/2024**:
   - *Issue*: Modern Excel defined-name parsers reject names with leading underscores combined with R1C1 tokens (such as `_XLCH_Row` or `_XLCH_Col`), throwing `Run-time error 1004: The syntax of this name isn't correct`.
   - *Mitigation*: Defined name constants use the clean `XLCH_` prefix (`XLCH_Row`, `XLCH_RowEnd`, `XLCH_Col`, `XLCH_ColEnd`, `XLCH_Excluded`, `XLCH_SheetExcluded`).

3. **VBA "Break on All Errors" Mode Compatibility**:
   - *Issue*: If Excel's VBA Editor option is set to *Break on All Errors* (*Tools -> Options -> General*), direct key lookup like `wb.Names("non_existent_name")` halts execution even under `On Error Resume Next`.
   - *Mitigation*: All defined-name lookups utilize `GetNameObject` safe collection iteration (`For Each n In wb.Names`), guaranteeing zero unhandled exceptions regardless of VBA IDE error trapping settings.

4. **Bounded Range Optimization**:
   - *Issue*: The highlight range is deliberately bounded to `UsedRange ∪ VisibleRange` to prevent performance degradation on 1,048,576-row spreadsheets.
   - *Behavior*: On sheets with sparse data separated by thousands of empty rows, scrolling deep into unpopulated areas might not show the highlight until you select a cell or interact with that area.

5. **Protected Worksheets**:
   - *Behavior*: Worksheets protected without the "Allow formatting cells" permission cannot accept conditional formatting overlay rules unless **Allow Protected** mode is toggled ON in the add-in options.

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
