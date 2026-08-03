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
- **Real colour swatch previews** in every gallery and on the Custom Colour
  button - the swatches are generated safely in C# (the COM ribbon add-in),
  so they render correctly on 64-bit Excel 2024/365 (the old VBA GDI
  approach crashed those builds).
- **Per-mode colour profiles** - in Crosshair mode, row and column highlights can use different colours.
- **Per-workbook or all-workbooks scope** - the highlight toggle can affect
  only the workbook you are in (default), or every open workbook at once
  (the **All Workbooks** ribbon toggle). See the scope section below.
- **Toggle from anywhere** - the **Highlight: On / Off** ribbon button, the
  **Ctrl+Shift+H** hotkey, and a **Toggle Highlighter** item in Excel's
  cell right-click menu all do the same thing, so you never have to leave
  your keyboard (or your context menu) to switch the highlight.
- **Configurable hotkeys** - toggle, history back, and history forward keys can be customised via the registry.
- Full merged-cell support: highlights automatically cover the entire dimensions of merged ranges.
- Dark mode tinting for pleasant visual contrast on dark Office themes.
- Workbook and sheet exclusion toggles stored in file-level defined names.
- Selection history navigation (default `Ctrl+Shift+Z` back, `Ctrl+Shift+X` forward).
- Named profiles for quick configuration switching.
- Settings persist between Excel sessions.
- No VBA project password, no hidden macros required in your own workbooks.
- The add-in's own VBA project is named **Highlighter** (you'll see
  `Highlighter (excel-highlighter.xlam)` in the VBE Project Explorer, not
  the default `VBAProject`).
- Designed to degrade gracefully on protected sheets and chart sheets rather than throwing errors at you.

## How it works, briefly

Four hidden, workbook-scoped defined names (`XLCH_Row`, `XLCH_RowEnd`, `XLCH_Col`, `XLCH_ColEnd`) hold the active selection bounds as plain numbers. Conditional formatting rules on each worksheet reference those names (`=AND(ROW()>=XLCH_Row,ROW()<=XLCH_RowEnd)`, etc.). Moving the active cell or selecting merged ranges updates the four names - a cheap operation - rather than adding or removing formatting rules on every keystroke. The CF rules themselves are only (re)built the first time a sheet is visited, or when you change mode, colour, style, or options from the ribbon.

See [docs/architecture.md](docs/architecture.md) for the full picture, including why the highlight range is bounded rather than applied to entire 1,048,576-row sheets.

## User Guide & Documentation

For a comprehensive step-by-step feature reference, shortcuts, and usage examples, see [docs/user-guide.md](docs/user-guide.md).

Quick reference:
- **Toggle Highlight**: Click **Highlight: On / Off** on the Ribbon, press
  `Ctrl+Shift+H` (configurable), or right-click any cell and choose
  **Toggle Highlighter** from the context menu. All three do exactly the
  same thing, and all three respect the per-workbook / all-workbooks scope
  setting.
- **Mode Selection**: Click **Row**, **Column**, **Crosshair**, or **Cell**.
- **Per-Mode Colours**: Toggle **Per-Mode Colours** in the Appearance group, then pick separate colours for row and column.
- **Colour Swatches**: The Colour, Row Colour, and Column Colour galleries show
  real colour swatches, and the Custom Colour button shows the current
  highlight colour.
- **Scope**: By default the highlight toggle only affects the workbook you
  are in. Switch the **All Workbooks** toggle in Options to apply it to every
  open workbook at once (and turn it off from any one of them).
- **Navigate Cell History**: Press `Ctrl+Shift+Z` (back) or `Ctrl+Shift+X` (forward) - both configurable.
- **Dark Mode**: Click **Dark Mode** under Options for comfortable dark-theme reading.

## Modes, styles & colours, explained

**The four modes** are the shape of the highlight:

- **Row** - highlights the whole row of the active cell.
- **Column** - highlights the whole column of the active cell.
- **Crosshair** - both at once, so your eye can trace the active row *and* column.
- **Cell** - highlights only the active cell (or merged range).

**Fill vs Border** (Style group) controls *how* the highlight is drawn:

- **Fill** (default) - tints the cells' interior colour.
- **Border** - draws a thick outline around the highlighted cells instead of
  tinting them, which is handy when you need to read dense data without
  obscuring any background formatting. It works in every mode, including
  Column and Crosshair.

**Intersect** (Style group, Crosshair mode) - the **intersection cell** is
where the highlighted row and column cross - the cell your cursor is actually
in. When **Intersect** is on, that cell gets an extra accent so it stands out
from the rest of the crosshair. The accent is drawn as a *separate* rule with
**StopIfTrue**, so it always wins at the cursor cell (older builds let the
row/column rules paint over it, which is why it looked like it did nothing),
and it is deliberately **darker** than the row/column colour so it is always
visible against them. With it off, the row and column are tinted uniformly
and the active cell is only marked by Excel's normal selection border. In
Cell mode the intersect accent *is* the highlight. In Border style the
accent also fills the cell so your cursor position stays obvious.

**Colours** - the seven presets are fixed RGB values. Prior to 2.2.0 the
palette was stored in the wrong byte order, so "Yellow" rendered as cyan and
"Blue" as violet. The palette values are now correct (Yellow =
`RGB(255,255,0)`, Blue = `RGB(153,204,255)`, ...), the Dark Mode tint and the
Pulse animation no longer swap red and blue, and the galleries show real
swatches so what you see is what you get.

**Pulse** (Style group) - when enabled, moving the active cell briefly
flashes the highlight at the new position (a quick lighter-then-back 3-step
cycle). Older builds scheduled the animation with a rounded-down timestamp
so the pulse often never visibly fired; it now uses full-precision timing and
stale animation chains self-cancel, so the flash reliably plays and always
restores the exact configured colour.

**Default state** - a fresh install (or **Reset Settings**) starts with the
highlight **off**, in **Row** mode, coloured **true yellow** `RGB(255,255,0)`
(fill style), per-workbook scope.

## Scope: one workbook or all of them

The **All Workbooks** toggle (Options group) decides what the **Highlight: On
/ Off** button affects:

- **This Workbook (default)** - toggling the highlight on only turns it on in
  the workbook you are currently in. Other open workbooks stay untouched.
  Turn it off in that workbook and the others keep their highlight.
  Ideal when you work across several files at once and only want the
  crosshair in one of them.
- **All Workbooks** - toggling on turns the highlight on for *every* open
  workbook, and toggling it off from any one of them turns it off for all of
  them.

The scope choice is remembered between sessions. Use **Reset Settings** to
return to the per-workbook default.

## Installation

### Option 1 - 1-click installer (recommended)

Double-click `install.bat` (or run `install.ps1` in PowerShell). The installer:

1. builds `excel-highlighter.xlam` from the source in `src/` and `customUI/`,
2. compiles and registers the **COM ribbon add-in** (`ExcelHighlighterRibbon.dll`)
   which delivers the Highlighter tab on Excel 2024 / recent Microsoft 365
   (see "Why is there a COM add-in?" below),
3. copies the add-in and icons to `%APPDATA%\Microsoft\AddIns\ExcelHighlighter`,
4. activates the add-in in the registry (`HKCU\...\Excel\Options`), and
5. clears Excel's ribbon cache and any Disabled Items entries so the
   Highlighter tab appears immediately.

### Option 2 - manual (pre-built .xlam)

See [docs/installation.md](docs/installation.md) for the step-by-step
instructions (create the add-ins folder, copy the .xlam + `images/`, tick it
in Excel Options > Add-ins).

### Option 3 - build from source

See the [Building from source](#building-from-source) section below.

### Uninstalling

Double-click `uninstall.bat` (or run `uninstall.ps1`). It removes the registry
activation, the COM add-in registration, and the deployed files.

## Building from source

The VBA project lives inside `excel-highlighter.xlam` (a binary Office file),
so the source of truth is the text-exported modules under `src/` plus the
RibbonX definition under `customUI/`.

### Prerequisites

- Windows with **Excel 2010 or newer** installed (the build drives Excel via
  COM automation).
- **.NET Framework 4.x** (present on every Windows 10/11) - needed only for
  the COM ribbon add-in build.
- **"Trust access to the VBA project object model"** enabled in Excel:
  File → Options → Trust Center → Trust Center Settings → Macro Settings.
  The build script imports modules into a fresh .xlam through the VBA
  Extensibility object model, which requires this setting.

### Build the .xlam

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build-xlam.ps1
```

This creates a blank `.xlam`, imports every module from `src/`, renames the
VBA project to `Highlighter`, injects the `customUI14.xml` ribbon part
(plus content-type and root/workbook relationships), and verifies the
resulting package. It finishes with a compile probe that opens the built
file in a hidden Excel instance and runs one `Application.Run` call per
module - a quick smoke test that the project compiles.

The build is idempotent: re-running overwrites `excel-highlighter.xlam`.

### Build the COM ribbon add-in

```powershell
powershell -ExecutionPolicy Bypass -File scripts\build-comribbon.ps1
```

Compiles `comribbon/Connect.cs` into `ExcelHighlighterRibbon.dll` with the
.NET Framework `csc.exe`, deploys it next to the installed .xlam, and
registers it in HKCU (no admin needed). This is a separate step because it
only matters on Excel 2024 / recent 365 builds - see below.

`install.ps1` runs both of these automatically.

### Why is there a COM add-in?

Empirically established on Excel 2024 (build 16.0.20228): Excel **never
queries the VBA ribbon path for an `.xlam`** - not the customUI package part,
not `IRibbonExtensibility.GetCustomUI` on ThisWorkbook. The add-in loads,
macros run, highlighting works, but no ribbon tab ever appears and no error
is shown. Excel *does* query a registered **COM add-in** for
`IRibbonExtensibility` on every build, so the Highlighter tab is delivered by
a small COM add-in (`comribbon/Connect.cs` → `ExcelHighlighterRibbon.dll`)
that returns the same `customUI14.xml` and forwards every ribbon callback
back into the VBA add-in via `Application.Run`. The .xlam's own customUI part
and VBA interface are kept as the classic mechanism for older builds.

If you build only the .xlam and install it manually on Excel 2024/recent 365,
you will get the highlighting but **no ribbon tab** - run `install.ps1` (or
`build-comribbon.ps1` + register) so the COM add-in is present too.

### Development loop (edit → export → commit)

1. Open the `.xlam` in Excel, make your changes in the VBA editor (Alt+F11).
2. Run `scripts/export-vba.ps1` to write the current VBA project back out to `src/` as text.
3. Commit the exported text files.

Going the other direction (text → binary) is exactly the build described
above. Both directions need Excel + "Trust access to the VBA project object
model" enabled.

### Signing (optional but recommended for the Publisher column)

Excel shows the **Publisher** column in the Add-ins dialog from the *digital
signature of the VBA project* - not from document properties (Author/Company)
and not from an Authenticode file signature. An unsigned add-in shows a blank
Publisher. Signing does **not** change any add-in behaviour, and it is *not*
required for the ribbon (see the note in docs/installation.md and the
2.1.3/2.1.4 changelog entries) - it only makes the add-in look properly
published in the Add-ins dialog and avoids an "unverified publisher" prompt
on machines that trust the certificate.

The VBA object model exposes **no signing API**, so the actual signing is a
one-time manual click-through in the Visual Basic Editor. `scripts/sign-xlam.ps1`
automates everything around it:

```powershell
# sign the repo build, deploy it over the installed copy, and trust the cert
powershell -ExecutionPolicy Bypass -File scripts\sign-xlam.ps1 -Rebuild -Deploy -Trust
```

What the script does:

1. Creates a self-signed code-signing certificate named **Sandeep Khadka** in
   your Personal certificate store if one doesn't exist yet.
2. Ensures the VBA signing defaults exist in the registry - `V1HashEnhanced=2`
   (SHA-256) and an RFC 3161 `TimeStampURL` under
   `HKCU\Software\Microsoft\VBA\Security`. This is the documented fix for
   the failure mode where signature parts exist but Excel's verdict is
   `Signed=False` (the decoded signature's digest algorithm reads empty);
   without it the signature consistently comes back invalid. The timestamp
   also lets the signature survive certificate expiry.
3. Pre-flights the machine: it refuses to start while a **visible** Excel
   window is open, and it detects **windowless zombie `EXCEL.EXE` processes**
   (leftovers from COM automation - no window, no tray icon, but they still
   lock the add-in file), reports their PIDs and offers to kill them, so the
   script never dead-ends on "Excel is currently running" when no Excel is
   visible anywhere.
4. Closes/holds Excel, opens the .xlam, and shows the VBA editor. It closes
   any copy of the add-in that auto-loaded via the registry first, so the
   VBE shows exactly **one** project - otherwise two projects both named
   `Highlighter (excel-highlighter.xlam)` appear and you can sign the wrong
   one.
5. **Waits for you** to do the one manual step (see below), then verifies the
   result with Excel's own verdict (`VBProject.Signed` on a fresh open).
6. Re-injects the customUI ribbon wiring that Excel's re-save strips
   (a signed-but-unwired file is worse than unsigned - see the changelog).
7. Locks the file read-only so Excel cannot re-save it and invalidate the
   signature again. Use `-NoLock` to skip.

The manual step, inside the VBA editor that the script opens:

1. Select `Highlighter (excel-highlighter.xlam)` in the Project Explorer (Ctrl+R if hidden).
2. Tools → Digital Signature... (click **Remove** first if a stale signature is shown).
3. Choose... → select **Sandeep Khadka** → OK → OK.
4. **Save the add-in yourself with Ctrl+S** inside the VBA editor (the
   script's own COM save can crash with `RPC_E_DISCONNECTED`; the native VBE
   save is reliable).
5. Switch back to the terminal and press Enter so the script can verify.

Useful switches:

- `-XlamPath <path>` - sign a specific copy instead of the default.
- `-Deploy` - copy the signed file over the installed copy afterwards.
- `-Trust` - also import the certificate into Trusted Publishers / Trusted
  Root so Excel doesn't warn that the publisher is unverified (Trusted Root
  needs admin; the Publisher column shows the name either way).
- `-Rebuild` - build from source first, then sign (build → sign → deploy).
  The sign target defaults to the **freshly built repo copy**, not the stale
  installed copy, so the new build actually ships.
- `-NoLock` - don't set the read-only lock (not recommended - the lock is
  what stops Excel from breaking the signature).

Notes:

- Re-running `install.bat` or `build-xlam.ps1` rebuilds the add-in and
  **removes the signature** - re-run `sign-xlam.ps1` afterwards.
- Self-signed certificates expire after a year. When the signature stops
  being trusted, delete the old certificate and re-run the script.
- The release builds on GitHub are shipped **unsigned** because signing is a
  per-machine, interactive step (each machine needs its own trusted cert).
  Sign locally after installing if you want the Publisher column filled in.

## Troubleshooting: ribbon tab missing / VBA errors

**The confirmed root cause (2.1.4) - and the fix:**

1. **Excel 2024 / recent Microsoft 365 never queries the VBA ribbon path
   for an `.xlam`.** Empirically proven on Excel 2024 build 16.0.20228: the
   add-in loads and runs but Excel silently never calls
   `IRibbonExtensibility.GetCustomUI` on ThisWorkbook and ignores the
   customUI package part - so the Highlighter tab never appears, with no
   error, on any delivery mechanism the VBA add-in can offer. (Earlier
   rounds blamed the digital signature and a "password-locked project" -
   both were overturned by decisive experiments.)
2. **The fix: a small COM add-in delivers the ribbon.** Excel DOES query a
   registered COM add-in for `IRibbonExtensibility` on every build.
   `install.ps1` (2.1.4+) automatically compiles and registers
   `ExcelHighlighterRibbon.dll` (source: `comribbon/Connect.cs`), which
   returns the very same `customUI14.xml` and forwards every ribbon
   callback back into the VBA add-in via `Application.Run` (see the
   `*Value` bridge functions in `RibbonCallbacks.bas`). **Re-run
   `install.bat` (or `install.ps1`) after closing Excel** - no manual step
   needed.
3. **The XML had latent schema bugs that could never surface until Excel
   parsed it (fixed in 2.1.4).** Once the COM add-in got the XML parsed,
   two real violations showed up and blocked the tab: `label` and
   `getLabel` declared together on the two exclude buttons (mutually
   exclusive), and several `imageMso` IDs that don't exist on Excel 2024.
   Both are fixed in `customUI/customUI14.xml`.
4. **Defined names used a prefix modern Excel rejects (fixed in 2.1.4).**
   Names like `_XLCH_Row` throw "The syntax of this name isn't correct"
   (#1004) on Excel 2024, which silently killed highlighting. All names now
   use the clean `XLCH_` prefix (`XLCH_Row`, `XLCH_Col`, ...).

If you've covered the above and the tab is still missing:

5. **Re-run `install.bat` after fully closing Excel.** The installer needs
   Excel closed to clear its ribbon cache, copy the file, and register the
   COM add-in.
6. **Check `File > Options > Add-ins > Manage: Disabled Items > Go...`.**
   If the add-in ever errored while loading, Excel silently blacklists it
   here and won't try loading it again until you remove it. `install.ps1`
   clears entries referencing this add-in automatically.
7. **Check `File > Options > Add-ins > Manage: COM Add-ins > Go...`.**
   Ensure **"Excel Highlighter Ribbon"** is checked - this COM add-in is
   what renders the tab on Excel 2024 / recent 365 builds.
8. **Check `File > Options > Trust Center > Trust Center Settings > Macro
   Settings`.** The ribbon buttons are VBA macros; if macros are set to
   "Disable all macros without notification," the tab may render but every
   button will fail silently.
9. **Ribbon button shows a stale "off" after toggling from the right-click
   menu or hotkey?** Fixed in 2.3.1: the COM add-in keeps the live ribbon
   in a static field and VBA invalidates it through
   `CreateObject("ExcelHighlighter.Ribbon").InvalidateRibbon()` after every
   external toggle. If it still doesn't refresh, check
   `%TEMP%\eh-com-invalidate.txt` for `InvalidateRibbon OK` vs
   `static _ribbon is NULL`.
10. Still stuck? Check the log at
   `%APPDATA%\ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log` for
   the actual VBA error and open an issue with it.

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
├── comribbon/
│   └── Connect.cs             COM ribbon add-in (Excel 2024 / recent 365)
├── customUI/
│   ├── customUI14.xml         RibbonX definition (Office 2010+ / customUI14)
│   └── images/                colour swatch PNGs (legacy - ribbon uses C# swatches)
├── docs/
│   ├── user-guide.md          complete feature & usage reference guide
│   ├── installation.md        installation & troubleshooting guide
│   ├── architecture.md        technical architecture & performance notes
│   └── future-features.md     project roadmap & feature backlog
├── scripts/                   PowerShell helpers
│   ├── build-xlam.ps1         text → .xlam build (Excel COM automation)
│   ├── build-comribbon.ps1    C# → ExcelHighlighterRibbon.dll build + register
│   ├── export-vba.ps1         .xlam → text export
│   ├── import-vba.ps1         text → .xlam import (low-level)
│   ├── sign-xlam.ps1          guided digital-signing + verify + lock
│   ├── lock-vba-project.ps1   optional password lock (see 2.1.3 note)
│   └── xlam-ribbon-utils.ps1  shared OPC zip-surgery helpers
├── tests/
│   └── manual-test-checklist.md
├── CHANGELOG.md
├── LICENSE
└── README.md
```

## Toggling the highlighter

Every toggle path - the **Highlight: On/Off** ribbon button, the
**Ctrl+Shift+H** hotkey, and the **Toggle Highlighter** item in the cell
right-click menu - goes through the same scope-aware toggle, and each one
keeps the ribbon button's pressed state and label in sync with the workbook
you toggled (the COM ribbon add-in holds its own ribbon reference and
invalidates it on every state change, so toggling from the context menu or
hotkey immediately updates the ribbon - no restart needed).

## Known limitations

- The highlight is bounded to each sheet's used range plus whatever is
  currently visible on screen, not the entire 1,048,576 × 16,384 grid. On a
  sheet with very sparse, widely separated data, scrolling into a completely
  empty area far from both isn't highlighted until you interact with it. See
  `docs/architecture.md` for the reasoning.
- Conditional formatting rules count against Excel's per-sheet CF limit.
  This add-in only ever adds one to three rules per sheet, but if a workbook
  already has an unusually large number of existing rules from other
  sources, adding ours could push it toward that ceiling.
- Sheets protected without "Allow formatting cells" enabled won't be
  highlighted unless "Allow Protected" mode is enabled in the Options group.
- The COM ribbon add-in is a .NET Framework assembly registered in HKCU. It
  requires .NET Framework 4.x (present on Windows 10/11) and is not needed
  on Excel builds that honour the .xlam's own ribbon part.

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

**Why is the Publisher column blank in the Add-ins dialog?**
The add-in is unsigned (or its signature was invalidated by an Excel re-save
- the VBE's Tools > Digital Signature dialog can still show the old name in
that case). See the [Signing](#signing-optional-but-recommended-for-the-publisher-column)
section - it's a one-time manual step (`scripts/sign-xlam.ps1`).

**Why does the signing script say "Excel is currently running" when no
Excel is open?**
A windowless zombie `EXCEL.EXE` (leftover from COM automation) is holding
 the add-in file - it has no window or tray icon to close. `sign-xlam.ps1`
 now detects these, shows their PID, and offers to kill them before
 continuing.

**Why do I need a COM add-in? I just want the highlighting.**
On Excel 2024 / recent 365 the highlighting works from the .xlam alone, but
the ribbon tab is only rendered by the COM add-in. Older Excel versions
don't need it. `install.ps1` handles both automatically.

## Contributing

Issues and pull requests are welcome. If you're proposing a new ribbon
control or a new highlight mode, please open an issue first to talk through
the approach - the CF-overlay design has some non-obvious constraints (see
architecture.md) that are worth checking against before writing code.

## License

MIT - see [LICENSE](LICENSE).
