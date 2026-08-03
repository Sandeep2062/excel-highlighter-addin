# Installation Guide

This guide covers every way to install the Excel Highlighter add-in, from the
one-click installer to building from source, plus the optional signing step
and a troubleshooting section.

## Before you start

- **Windows** with **Excel 2010 or newer** (the add-in is tested on Excel 2024
  and recent Microsoft 365, and works on older builds).
- **Close Excel before installing.** A running Excel instance locks the add-in
  file and can hold the old copy in memory.
- The installer needs no administrator rights (everything lives under your
  user profile).

---

## Option A - 1-Click Automated Installer (Recommended)

1. Clone or download this repository.
2. Double-click `install.bat` (or right-click `install.ps1` → **Run with
   PowerShell**).
3. The script will automatically:
   - Build `excel-highlighter.xlam` from source (`src/` + `customUI/`).
   - Compile and register the **COM ribbon add-in**
     (`ExcelHighlighterRibbon.dll`), which is what renders the Highlighter
     tab on Excel 2024 / recent Microsoft 365.
   - Copy `excel-highlighter.xlam` and the `images/` folder to
     `%APPDATA%\Microsoft\AddIns\ExcelHighlighter\`.
   - Activate the add-in in the registry
     (`HKCU:\Software\Microsoft\Office\<version>\Excel\Options`).
   - Clear Excel's ribbon cache and any **Disabled Items** entries.
4. Open Microsoft Excel — the **Highlighter** tab appears on your ribbon.

> [!NOTE]
> To cleanly uninstall, double-click `uninstall.bat` (or run `uninstall.ps1`).

---

## Option B - Manual Installation (Pre-built .xlam)

Use the pre-built `excel-highlighter.xlam` + `images/` folder from a release.

1. Create the folder `%APPDATA%\Microsoft\AddIns\ExcelHighlighter\`.
2. Copy `excel-highlighter.xlam` **and** the whole `images/` folder inside it.
3. Open Excel: **File → Options → Add-ins → Manage: Excel Add-ins → Go... →
   Browse...**.
4. Select the `.xlam` and make sure its checkbox is ticked, then **OK**.

> [!IMPORTANT]
> On **Excel 2024 / recent Microsoft 365** the ribbon tab is delivered by the
> **COM add-in** (`ExcelHighlighterRibbon.dll`), not by the `.xlam`'s own
> ribbon part (see the 2.1.4 changelog entry for the evidence). If you only
> copy the `.xlam` manually, highlighting works but **no ribbon tab appears**.
> Either run `install.bat` (which registers the COM add-in for you) or run
> `scripts/build-comribbon.ps1` to compile and register it.

---

## Option C - Build from Source

### Prerequisites

- Windows with **Excel 2010+** installed.
- **"Trust access to the VBA project object model"** enabled:
  **File → Options → Trust Center → Trust Center Settings → Macro Settings →
  tick "Trust access to the VBA project object model"**. The build imports
  modules through the VBA Extensibility object model, which requires this.
- **.NET Framework 4.x** (present on Windows 10/11) — only needed for the COM
  ribbon add-in step.

### Steps

1. From PowerShell, build the `.xlam`:

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\build-xlam.ps1
   ```

   This creates a blank `.xlam`, imports every module from `src/`, renames the
   VBA project to **Highlighter**, injects the `customUI14.xml` ribbon part
   (plus content-type and root/workbook relationships), verifies the package,
   and finishes with a compile probe (10 `Application.Run` checks across all
   modules). The build is idempotent — re-running overwrites
   `excel-highlighter.xlam`.

2. Install it. Either run `install.bat`, or follow **Option B** with the
   freshly built file.

3. **Only on Excel 2024 / recent 365**: build and register the COM ribbon
   add-in so the tab appears:

   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\build-comribbon.ps1
   ```

   This compiles `comribbon/Connect.cs` into `ExcelHighlighterRibbon.dll`
   with the .NET Framework `csc.exe`, deploys it next to the installed .xlam,
   and registers it under HKCU (no admin needed). `install.bat` runs both
   build steps automatically.

---

## Option D - Setting the Publisher name (digital signature)

Excel shows the **Publisher** column in the Add-ins dialog from the *digital
signature of the VBA project* — document properties (Author/Company) and
Authenticode file signatures do **not** affect it. An unsigned add-in shows a
blank Publisher. Signing is **optional**: it does not change add-in behaviour
and is not required for the ribbon — it only fills in the Publisher column.

The VBA object model exposes **no signing API**, so the actual signing is a
one-time manual click-through in the Visual Basic Editor.
`scripts/sign-xlam.ps1` automates everything around it:

```powershell
# sign the repo build, deploy it over the installed copy, and trust the cert
powershell -ExecutionPolicy Bypass -File scripts\sign-xlam.ps1 -Rebuild -Deploy -Trust
```

What the script does, step by step:

1. Creates a self-signed code-signing certificate named **Sandeep Khadka** in
   your Personal certificate store if one doesn't exist yet.
2. Ensures the VBA signing defaults exist in the registry —
   `V1HashEnhanced=2` (SHA-256) and an RFC 3161 `TimeStampURL` (DigiCert)
   under `HKCU\Software\Microsoft\VBA\Security`. Without `V1HashEnhanced`,
   Excel consistently reports the signed file as **Stale** (signature parts
   exist but its own fresh-open verdict is `Signed=False`) — the signature
   never sticks and the Publisher column stays blank. Existing values are
   left untouched.
3. Pre-flight checks: refuses if a **visible** Excel window is open, and if it
   finds **windowless zombie `EXCEL.EXE` processes** (leftovers from COM
   automation — no window, no tray icon, but they still lock the add-in file)
   it reports their PID and offers to kill them. With `-Rebuild` the fresh
   build in the repo root is the sign target (not the stale installed copy).
4. Strips any leftover signature parts so every attempt starts from a clean
   slate, then opens the add-in in Excel and shows the Visual Basic Editor.
5. **Waits for you** to do the one manual step (below), then verifies the
   result with Excel's own verdict (`VBProject.Signed` on a fresh open).
6. Re-injects the customUI ribbon wiring that Excel's re-save strips (a
   signed-but-unwired file is worse than unsigned).
7. Locks the file read-only so Excel cannot re-save it and invalidate the
   signature again. Use `-NoLock` to skip.

The manual step, inside the VBA editor the script opens:

1. Select `Highlighter (excel-highlighter.xlam)` in the Project Explorer
   (Ctrl+R if hidden). The script closes any auto-loaded copy first, so there
   is exactly **one** project with that name.
2. **Tools → Digital Signature...** — click **Remove** first if a stale
   signature is shown.
3. **Choose... → select Sandeep Khadka → OK → OK**.
4. **Save the add-in yourself with Ctrl+S** inside the VBA editor (the
   script's own COM save can crash with `RPC_E_DISCONNECTED`; the native VBE
   save is reliable).
5. Switch back to the terminal and press Enter so the script can verify.

Useful switches:

- `-XlamPath <path>` — sign a specific copy instead of the default.
- `-Deploy` — copy the signed file over the installed copy afterwards.
- `-Trust` — also import the certificate into Trusted Publishers / Trusted
  Root so Excel doesn't warn that the publisher is unverified (Trusted Root
  needs admin; the Publisher column shows the name either way).
- `-Rebuild` — build from source first, then sign (build → sign → deploy).
- `-NoLock` — skip the read-only lock (not recommended — the lock is what
  stops Excel from breaking the signature).

> [!NOTE]
> Re-running `install.bat` or `build-xlam.ps1` rebuilds the add-in, which
> **removes** the signature — run `sign-xlam.ps1` again afterwards.
> Self-signed certificates expire after a year; when the signature stops
> being trusted, delete the old certificate and re-run the script.
> The release builds on GitHub are shipped **unsigned** because signing is a
> per-machine, interactive step.

---

## Uninstalling / Disabling

- **Automated**: double-click `uninstall.bat` (or run `uninstall.ps1`). It
  removes the registry activation, the COM add-in registration, and the
  deployed files.
- **Manual**: in Excel, untick the add-in under
  **File → Options → Add-ins → Manage: Excel Add-ins**.

---

## Troubleshooting

- **No "Highlighter" tab appears, but the add-in loads (macros run).**
  On Excel 2024 / recent Microsoft 365 the ribbon is delivered by the **COM
  add-in**, not by the `.xlam`'s own customUI part (2.1.4 evidence). Fix:
  re-run `install.bat` after closing Excel — it compiles and registers
  `ExcelHighlighterRibbon.dll`. Then check
  **File → Options → Add-ins → Manage: COM Add-ins → Go...** and make sure
  **"Excel Highlighter Ribbon"** is ticked.
- **Add-in shows as installed/checked but never loads at all.** Older
  installers wrote the Excel `OPEN` registry value with a `/R` prefix, which
  silently prevents the add-in from loading (fixed in 2.1.0). Re-run
  `install.bat` after closing Excel — it sweeps stale `OPEN*` values.
- **Add-in stopped loading after an error.** Excel blacklists crashed
  add-ins in `HKCU\...\Resiliency\DisabledItems`. `install.ps1` clears
  entries referencing this add-in automatically; you can also remove them by
  hand.
- **The script says "Excel is currently running" but no Excel is open.**
  A windowless zombie `EXCEL.EXE` (leftover from COM automation) is holding
  the file. `sign-xlam.ps1` now detects these, reports their PID, and offers
  to kill them — press `y` at the prompt (or end them in Task Manager >
  Details).
- **Publisher column is blank in the Add-ins dialog.** The VBA project is
  not validly signed — run `scripts/sign-xlam.ps1` (Option D). Note that the
  VBE's **Tools → Digital Signature** dialog can show a certificate name even
  when Excel considers the signature invalid (it was invalidated by a
  re-save); only Excel's own fresh-open verdict counts.
- **Signed but Publisher is blank again later.** The add-in was re-saved by
  Excel (e.g. the "Remove personal information from file properties on save"
  privacy option under **File → Options → Trust Center → Trust Center
  Settings → Privacy Options**), which invalidates the signature. Re-run
  `sign-xlam.ps1`, which locks the file read-only so this cannot recur.
- **Colour gallery shows no swatch previews.** Since 2.2.0 the swatches are
  generated safely in C# by the COM ribbon add-in. If they don't render, the
  COM add-in isn't loaded — re-run `install.bat`.
- **Macros disabled.** The ribbon buttons are VBA macros; if Excel is set to
  "Disable all macros without notification", the tab may render but every
  button fails silently. Check **File → Options → Trust Center → Trust Center
  Settings → Macro Settings**.
- **Still stuck?** Check the log at
  `%APPDATA%\ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log` for the
  actual VBA error and open an issue with it.
