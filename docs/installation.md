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

## Option D - Setting the Publisher name (digital signature)

Excel shows the **Publisher** column in the Add-ins dialog from the *digital signature of the VBA project* - document properties (Author/Company) and Authenticode file signatures do **not** affect it. If the add-in is unsigned, Publisher is blank, exactly as shown in the Add-ins dialog.

To make it read **Sandeep Khadka**:

1. Close Excel (a running Excel instance locks the add-in file).
2. Run `scripts/sign-xlam.ps1` from PowerShell:
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts\sign-xlam.ps1
   ```
   The script will:
   - create a self-signed code-signing certificate named **Sandeep Khadka** in your personal certificate store if one doesn't already exist;
   - open the add-in in Excel and bring up the Visual Basic Editor;
   - wait while you sign the project once: **Tools → Digital Signature... → Choose... → Sandeep Khadka → OK → OK** (first single-click `VBAProject (excel-highlighter.xlam)` in the Project Explorer / Ctrl+R so you sign the right project; if the dialog already shows a signature, click **Remove** first so a stale one doesn't get reused);
   - have you **save the add-in yourself with Ctrl+S inside the VBA editor** (the script's own COM save can crash with `RPC_E_DISCONNECTED` if Excel is holding the file, so the script prefers the native VBE save);
   - **verify the signature cryptographically** - the script decodes the signature inside the saved file and checks that its digest matches the current project bytes (mere presence of the signature parts is not proof, since Excel silently invalidates the signature whenever it re-saves the project after signing);
   - **lock the add-in file read-only** so Excel cannot re-save it and invalidate the signature again (use `-NoLock` to skip the lock).

   > [!NOTE]
   > If the script prints "VBProject.Signed reports no signature", that warning is a known false negative of Excel's in-memory check - **ignore it**. Trust only the final `VERIFIED: the add-in file carries a VALID VBA digital signature...` line.
3. Restart Excel and open **File → Options → Add-ins** - the Publisher column now reads **Sandeep Khadka**.

Useful switches:

- `-XlamPath <path>` - sign a specific copy (e.g. the repo build).
- `-Deploy` - after signing, copy the signed file over the installed copy under `%APPDATA%\Microsoft\AddIns\ExcelHighlighter`.
- `-Trust` - also add the certificate to Trusted Publishers / Trusted Root so Excel doesn't warn that the publisher is unverified (Trusted Root needs administrator rights; the Publisher column shows the name either way).
- `-Rebuild` - rebuild from source first, then sign (build → sign → deploy).
- `-NoLock` - skip the read-only lock after signing (not recommended - the lock is what stops Excel from breaking the signature).

> [!NOTE]
> Re-running `install.bat` rebuilds the add-in, which **removes** the signature - run `sign-xlam.ps1` again afterwards. Self-signed certificates expire after a year; when the signature stops being trusted, delete the old certificate and re-run the script.

---

## Uninstalling / Disabling

- **Automated**: Double-click `uninstall.bat` (or run `uninstall.ps1`).
- **Manual**: In Excel, untick the add-in under **File → Options → Add-ins → Manage Excel Add-ins**.

---

## Troubleshooting

- **No "Highlighter" tab appears**: Two confirmed causes, in order of likelihood:
  1. **The add-in was registered with `/R` in the Excel `OPEN` registry value** (fixed in 2.1.0). Older `install.ps1` versions wrote `/R "...excel-highlighter.xlam"`. With `/R` the add-in shows as installed/checked in the Add-ins dialog but **never actually loads** - no `Workbook_Open`, no ribbon tab, no error. That is the exact symptom reported across multiple PCs and Excel versions. 2.1.0 writes the plain quoted path (the same format Excel and working add-ins such as ASAP Utilities use). If you installed with any earlier version, **re-run `install.bat` after closing Excel**. The installer now also sweeps and removes every stale `OPEN*` value referencing the add-in, not just the first one.
  2. **Excel 2024 / recent Microsoft 365 builds suppress the custom ribbon for unsigned add-ins (strongly indicated).** Verified empirically on Excel 2024 (build 16.0.20228): every delivery mechanism for an unsigned add-in - the customUI part, VBA `IRibbonExtensibility`, even a bare part-only add-in - fails to render its tab while macros still run. The one add-in whose ribbon does render on the same machine is digitally signed by a certificate in the user's Trusted Publishers store. Fix: `install.ps1` imports the self-signed "Sandeep Khadka" certificate into **Trusted Publishers** automatically, then run `scripts/sign-xlam.ps1 -Deploy -Trust` once (see Option D) to sign the add-in - the confirming test (signing and seeing the tab appear) is this manual step.

  Otherwise: ensure macros and add-ins are allowed in Excel Trust Center. If Excel was open during installation, restart Excel.
- **Colour gallery shows text labels instead of colour swatches**: expected - the ribbon deliberately uses native `imageMso` icons (see the 2.0.1 changelog note on why dynamic GDI swatches were removed to keep the tab loading reliably on 64-bit Office). The `images` folder is no longer referenced by the ribbon.
- **Errors log location**: Checked silently under `%APPDATA%\ExcelCrosshairHighlighter\ExcelCrosshairHighlighter.log`.
- **Publisher column is blank in the Add-ins dialog**: the VBA project is not validly signed - run `scripts/sign-xlam.ps1` (see Option D).
- **Script says "No signature detected" but I did sign**: the script's in-memory check is unreliable; trust the final cryptographic verification. If it says the signature is **stale** (parts present but digest mismatch), Excel re-saved the project after you signed it, invalidating the signature - the script's read-only lock prevents this from happening again; re-run it once more.
- **I signed but the Publisher is blank again after a while**: the add-in was re-saved by Excel (e.g. by the "Remove personal information from file properties on save" privacy option under **File → Options → Trust Center → Trust Center Settings → Privacy Options**). Re-run `sign-xlam.ps1`, which locks the file read-only so this cannot recur.
