# Installation Guide

## Option A - 1-Click Automated Installer (Recommended)

1. Clone or download this repository.
2. Double-click `install.bat` (or right-click `install.ps1` → **Run with PowerShell**).
3. The script will automatically:
   - Build `excel-highlighter.xlam` if not already present.
   - Copy `excel-highlighter.xlam` and `images/` icons to `%APPDATA%\Microsoft\AddIns\ExcelHighlighter\`.
   - Register and activate the add-in in Windows Registry (`HKCU:\Software\Microsoft\Office\<version>\Excel\Options`).
4. Open Microsoft Excel — the **Highlighter** tab will appear on your Ribbon automatically!

> [!NOTE]
> To cleanly uninstall anytime, double-click `uninstall.bat` (or run `uninstall.ps1`).

---

## Option B - Manual Installation (Pre-built .xlam)

1. Download `excel-highlighter.xlam` and the `images/` folder.
2. Create a folder `%APPDATA%\Microsoft\AddIns\ExcelHighlighter\` and place `excel-highlighter.xlam` and the `images` folder inside it.
3. Open Excel: **File → Options → Add-ins → Manage Excel Add-ins → Go... → Browse...**.
4. Select `%APPDATA%\Microsoft\AddIns\ExcelHighlighter\excel-highlighter.xlam` and make sure the checkbox is ticked.

---

## Option C - Manual Build from Source

Requirements: Excel 2010+ on Windows with "Trust access to the VBA project object model" enabled (**File → Options → Trust Center → Trust Center Settings → Macro Settings**).

1. Run `scripts/build-xlam.ps1` from PowerShell to build `excel-highlighter.xlam`.
2. Run `install.bat` or follow Option B to activate the add-in.

---

## Uninstalling / Disabling

- **Automated**: Double-click `uninstall.bat` (or run `uninstall.ps1`).
- **Manual**: In Excel, untick the add-in under **File → Options → Add-ins → Manage Excel Add-ins**.

---

## Troubleshooting

- **No "Highlighter" tab appears**: Ensure macros and add-ins are allowed in Excel Trust Center. If Excel was open during installation, restart Excel.
- **Colour swatches show as blank squares**: Confirm the `images` folder sits inside `%APPDATA%\Microsoft\AddIns\ExcelHighlighter\images\`.
- **Errors log location**: Checked silently under `%APPDATA%\ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log`.
