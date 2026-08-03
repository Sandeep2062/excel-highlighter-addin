# Changelog

All notable changes to this project are documented here.
Format loosely follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/).

## [2.4.0] - 2026-08-03

Release hardening pass: the signing workflow now survives the two failure
modes that previously dead-ended it (invisible zombie Excel processes, and
`-Rebuild -Deploy` signing the stale installed copy), plus a full docs
refresh (build + install guides) and a clean unsigned rebuild of the release
binary.

### Chores
- **Repo cleanup + docs refresh** - updated `README.md` with a detailed
  build & install guide, rewrote `docs/installation.md` (Option D signing
  flow now matches the script's real behaviour), corrected the stale
  `_XLCH_*` defined-name references in `docs/architecture.md` and
  `docs/future-features.md`, and rebuilt `excel-highlighter.xlam` fresh
  (unsigned, as release builds ship) so the committed binary matches the
  source and carries no leftover signature parts.
- **`APP_VERSION` bumped to 2.4.0**.

### Fixed
- **Signing consistently returned "Stale" (signature parts exist but Excel
  says `Signed=False`) - the signing script did not set the VBA signing
  defaults.** Session finding on Excel 2024 (build 16.0.20228): with no
  `V1HashEnhanced` registry value, the signature's digest algorithm decodes
  empty and Excel's own fresh-open verdict is `False` on a fully readable
  project. `scripts/sign-xlam.ps1` now ensures `V1HashEnhanced=2` (SHA-256)
  plus an RFC 3161 `TimeStampURL` (DigiCert) under
  `HKCU\Software\Microsoft\VBA\Security` before every signing attempt -
  the exact fix Microsoft's "Digitally sign your VBA macro project"
  documentation prescribes. Existing values are left untouched.
- **The VBE showed TWO projects both named `Highlighter
  (excel-highlighter.xlam)`, so the manual signing step could target the
  wrong one.** The add-in auto-loads via the `OPEN` registry value, so a
  fresh Excel instance already had a copy open before `sign-xlam.ps1`
  opened the file being signed - and the advisory check even reported the
  project as "not readable here". The interactive signing session now
  closes any auto-loaded copy first, so the VBE shows exactly one project
  with an unambiguous name.
- **`sign-xlam.ps1` refused to start with "Excel is currently running"
  even though no Excel window was visible anywhere.** COM automation
  (the `-Rebuild` build step and earlier aborted signing runs) leaves
  windowless zombie `EXCEL.EXE` processes behind - no window, no tray
  icon, hidden under Task Manager's Details/Background processes - yet
  they still hold the add-in file open and trip the old naive
  `Get-Process EXCEL` check. The pre-flight now distinguishes a visible
  Excel window (refused, as before) from windowless zombies: it reports
  their PID + start time and offers to kill them, after a short grace
  period for instances that are merely mid-shutdown. `build-xlam.ps1`
  now releases its workbook COM references and forces a GC before
  quitting, so the build stops leaving zombies behind in the first place.
- **`-Rebuild -Deploy` signed the STALE installed copy instead of the
  fresh build** - with no explicit `-XlamPath`, the default resolution
  preferred the installed copy under `%APPDATA%\Microsoft\AddIns`, so
  after a rebuild the script signed the OLD file and the deploy copy step
  was a no-op (source == target). The default now prefers the freshly
  built repo copy when `-Rebuild` is used, so `-Rebuild -Deploy` signs
  the new build and pushes it over the installed copy as documented.

## [2.3.1] - 2026-08-03

First tagged release (v2.3.1). Two user-reported polish fixes, both verified
empirically against the live add-in on Excel 2024 (interactive launch + UIA):

### Chores
- **Repo hygiene** - added `.gitignore` (Freebuff workspace, probe files,
  generated DLL), stopped tracking the `.freebuff/` scratch directory,
  removed probe workbooks / the no-part diagnostic build from the repo
  root, and committed the previously-untracked build essentials
  (`comribbon/Connect.cs`, `scripts/build-comribbon.ps1`,
  `scripts/xlam-ribbon-utils.ps1`, `scripts/lock-vba-project.ps1`) that a
  fresh clone needs to build the COM ribbon add-in.
- **`APP_VERSION` bumped to 2.3.1** (was stale at 2.2.0).
- **README rewritten** with a complete guide: features, usage, building
  from source (both the .xlam and the COM ribbon add-in), signing, and
  troubleshooting.

### Fixed
- **Ribbon toggle button showed stale "off" after toggling from the
  right-click menu or the hotkey** - the highlight turned on, but the
  ribbon's Highlight button never reflected it until Excel was restarted.
  Root cause (found by live diagnosis, not guessing): the ribbon is
  delivered by the `ExcelHighlighter.Ribbon` COM add-in, and BOTH VBA-side
  invalidation channels were broken on this build - (1) `Application.Run`
  cannot marshal an `IRibbonUI` argument, so `RibbonCallbacks.onLoad`
  throws and VBA's `mRibbon` stays Nothing (`RibbonAttached()`=False);
  (2) `Application.COMAddIns("...").Object` returns Nothing from VBA too.
  `InvalidateRibbon` therefore never reached the live ribbon. The COM
  add-in now keeps the ribbon it is handed in `OnLoad` in a **static
  field**, and VBA's `InvalidateRibbon` calls `CreateObject("ExcelHigh-
  lighter.Ribbon").InvalidateRibbon()` on a throwaway instance that shares
  that static - so the real ribbon gets `Invalidate()`d and Excel
  re-queries `getPressed`/`getLabel`. Verified live: the button label
  flips `Highlight: Off` → `Highlight: On` → `Highlight: Off` across
  context-menu toggles, log shows `CreateObject route=True`.

### Changed
- **VBA project renamed from "VBAProject" to "Highlighter"** - the VBE
  Project Explorer now shows `Highlighter (excel-highlighter.xlam)` instead
  of `VBAProject (excel-highlighter.xlam)`. `scripts/build-xlam.ps1` sets
  `VBProject.Name = "Highlighter"` after importing modules; safe because
  every runtime entry point (ribbon callbacks, `Application.Run`, `OnKey`,
  `OnTime`, context-menu `OnAction`) references procedures by
  module-qualified name only. `scripts/sign-xlam.ps1` and
  `scripts/lock-vba-project.ps1` instruction text updated to match.
- **`comribbon/Connect.cs`** - static `_ribbon` field + public
  `InvalidateRibbon()` method. (The 2.3.1 fix was originally written into
  `comribbon/Connect-full.cs`, a scratch copy that the build script never
  compiles - `scripts/build-comribbon.ps1` builds `comribbon/Connect.cs`,
  so the deployed DLL shipped WITHOUT the fix and the toggle stayed stale.
  The fix is now in the file the build actually compiles, and the deployed
  `ExcelHighlighterRibbon.dll` was rebuilt and re-registered.)
- **`src/RibbonCallbacks.bas`** - `InvalidateRibbon` now also invalidates
  via `CreateObject("ExcelHighlighter.Ribbon")`; added `RibbonAttached()`
  diagnostic.

## [2.3.0] - 2026-08-02

Fixes for the three things the user reported as still not working after
2.2.0 (Intersect, Pulse, and an easy toggle), plus a right-click toggle.
All changes verified empirically against the live add-in on Excel 2024 via
the direct-open COM harness: 14/14 checks passed.

### Fixed
- **Intersect accent looked like it did nothing - the row/column rules
  painted over it.** The accent was added as a plain rule with no
  priority; Excel's rule precedence let the later row/column rules win at
  the cursor cell, so the darker intersection tint never showed. The
  accent is now a dedicated rule with `StopIfTrue=True` and
  `SetFirstPriority`, so it always wins at the exact crossing cell, and it
  is deliberately **darkened** (`Utilities.DarkenColour`) so it's clearly
  visible against the row/column bands (empirically verified: accent
  `RGB(140,140,0)` vs row yellow `RGB(255,255,0)`). In Border style the
  accent also fills the cell, so the cursor cell stays obvious with a
  border-only highlight.
- **Pulse animation never visibly fired - the OnTime target was rounded
  to "hh:mm:ss"**, which dropped the date and sub-second precision, so
  Excel often received a time already in the past. `TriggerAnimationPulse`
  now schedules with a full-precision `Date` (`Now + 0.12/86400`), and a
  generation counter makes stale timer chains self-cancel instead of
  fighting newer pulses. Verified: the pulse chain starts on selection
  change, runs with zero errors, and restores the exact configured colour.

### Added
- **Right-click context-menu toggle** - a **Toggle Highlighter** item is
  added to Excel's standard Cell command bar (tagged `XLCH_Toggle`,
  `Temporary:=True`) on add-in start, checked/unchecked live from
  `EventApp.App_SheetBeforeRightClick`, and routed through the same
  scope-aware `RibbonCallbacks.SetHighlightForActiveWorkbook` toggle as
  the ribbon button and hotkey. Removed on unload.
- **Shortcut surfaced in tooltips** - the ribbon Toggle button and the
  history Back/Forward buttons now mention their hotkeys (e.g. `Ctrl+Shift+H`)
  in their screentips/supertips.

### Changed
- **`src/AddinHost.bas`** - context-menu add/remove/refresh + `OnContextToggle`
  handler. (Also fixed a latent compile error: `CONTEXT_BUTTON_TAG` was
  declared mid-module instead of in the declarations section, which made
  every procedure after it fail to compile with an untrappable error - the
  build's `IsRunning`-only compile check never caught it.)
- **`src/EventApp.cls`** - `SheetBeforeRightClick` hook calls
  `AddinHost.RefreshContextMenuState`.
- **`src/HighlightEngine.bas`** - `AddIntersectionRule` (StopIfTrue +
  first priority + border-style fill accent), generation-based pulse
  scheduling, pulse channel handling per-axis so per-mode colours survive
  mid-animation.
- **`src/Utilities.bas`** - new `DarkenColour(ByVal colour, ByVal percent)`
  helper (BGR-aware, returns a visibly darker tint).
- **`customUI/customUI14.xml` / `customUI.xml`** - hotkey hints in the
  Toggle and History tooltips.

## [2.2.0] - 2026-08-02

Refinement release (ribbon already rendering since 2.1.4). Every change below
was verified empirically against the live add-in on Excel 2024 (build
16.0.20228) via a direct-open COM harness that drives the real VBA engine:
10/10 checks passed (defaults, row/column border colour, per-workbook scope,
all-workbooks scope).

### Fixed
- **Preset colour palette was in the wrong byte order - "Yellow" rendered
  cyan and "Blue" rendered violet** - the RGB literals in `Constants.bas`
  were written as if they were `RRGGBB` hex, but Excel's Long colour format
  is `0x00BBGGRR` (blue in the high byte). `RGB_YELLOW = 16776960` is
  actually `RGB(0,255,255)` = cyan, and `RGB_BLUE = 16751001` is
  `RGB(153,153,255)` = violet. Corrected: Yellow = `RGB(255,255,0)` (65535),
  Blue = `RGB(153,204,255)` (16764057), and the other presets fixed to
  match their labels. The gallery labels now match what actually renders.
- **Dark Mode tint and Pulse animation swapped red and blue** - both
  `Settings.ApplyDarkModeTint` and `HighlightEngine.PulseStep` read the
  channels in RRGGBB order when the value is really BGR, so e.g. yellow
  turned cyan under Dark Mode. Both now decode `0x00BBGGRR` correctly.
- **Border style lost its colour (and weight) on Excel 2024 - the border
  rendered thin and black regardless of the chosen colour.** Root cause,
  found by inspecting the saved OOXML dxf: `FormatCondition.Borders` does
  NOT accept `.Weight = xlThick` (raises error 1004 on this build). The old
  `AddRule` set Weight *before* Color, so the error aborted the block before
  `.Color` was assigned. `AddRule` now sets Color first, then LineStyle,
  then Weight as a tolerated best-effort, so the border renders in the
  requested colour in every mode (Row, Column, Crosshair, Cell).
  (The 2.2.0 gallery swatches + OOXML verification also confirmed the
  border rule itself is created in Column mode - the old "border only
  worked in row" behaviour is gone.)
- **Colour galleries had no icon previews** - the galleries and the Custom
  Colour button now render real colour swatches, generated safely in C#
  (`Connect.MakeSwatch`, `GetGallery_ItemImage` / `GetSwatchImage`). This
  is the safe re-enable of the 2.0.1-removed GDI swatches: instead of VBA
  `OleCreatePictureIndirect` (which crashes 64-bit Office with
  `0x8000FFFF`), the swatches are built with `System.Drawing.Bitmap` in the
  COM add-in and handed to Excel as `IPictureDisp` via `AxHost`. Cached by
  RGB value so repaint-heavy polling doesn't rebuild them.

### Added
- **Per-workbook vs all-workbooks toggle scope** - a new **All Workbooks**
  toggle in the Options group (Settings.ScopeAll, default False):
  - *This Workbook (default)* - the Highlight On/Off toggle affects only the
    workbook you are in; other open workbooks stay untouched and keep their
    own state.
  - *All Workbooks* - toggling on turns the highlight on for every open
    workbook, and turning it off from any one workbook turns it off for all.
  - Implementation: `HighlightEngine.IsWorkbookHighlightActive` is the
    single gate; per-workbook state lives in a session-only dictionary
    (`SetWorkbookEnabled` / `WorkbookEnabled` / `ClearWorkbookEnabled`),
    and the hotkey + ribbon toggle both route through it.
- **Colour swatch previews** - `getItemImage` on all three colour galleries
  (Colour, Row Colour, Col Colour) and `getImage` on the Custom Colour
  button, backed by `RibbonCallbacks.GetGallery_ItemRGBValue` /
  `GetSwatchRGBValue` + the C# swatch generator.

### Changed
- **Default state is now Row mode + true yellow** - a fresh install or
  **Reset Settings** starts with the highlight off, in **Row** mode,
  coloured `RGB(255,255,0)` (fill), per-workbook scope. (The old defaults
  were inconsistent with the labels and relied on the broken palette.)
- **`src/Constants.bas`** - palette literals corrected (see Fixed),
  `CTRL_SCOPE_ALL` added, `APP_VERSION` bumped to 2.2.0.
- **`src/HighlightEngine.bas`** - scope gate, per-workbook state,
  `ReapplyAllOpenWorkbooks` handles per-workbook vs all-workbooks stripping,
  `AddRule` colour-first border fix, pulse channel-order fix.
- **`src/RibbonCallbacks.bas`** - scope toggle callbacks + `*Value` bridges,
  swatch RGB bridges.
- **`src/AddinHost.bas`** - hotkey toggle routes through the scope-aware
  `HighlightEngine` path.
- **`src/Profiles.bas`** - default mode saved/loaded as Row.
- **`comribbon/Connect.cs`** - `GetScopeAll_*`/`OnScopeAll_Action`,
  `GetGallery_ItemImage`, `GetSwatchImage`, `MakeSwatch` + swatch cache
  (System.Drawing + AxHost IPictureDisp helper).
- **`customUI/customUI14.xml` / `customUI.xml`** - new All Workbooks toggle,
  `getItemImage`/`getImage` swatch wiring, Reset tooltip updated.
- **`scripts/build-comribbon.ps1`** - references System.Drawing +
  System.Windows.Forms for swatch generation.
- **README + docs/user-guide.md** - new in-depth sections explaining the
  modes, Fill vs Border, **Intersect** (the row/column crossing cell), the
  corrected colour palette, the default state, and the per-workbook vs
  all-workbooks scope.

## [2.1.4] - 2026-08-02

**The ribbon now renders on Excel 2024 (build 16.0.20228) - verified end-to-end.**
The 2.1.3 "password-locked project is the gate" conclusion was also WRONG -
locking the project made no difference. The definitive finding this round:
**Excel 2024 never queries the VBA ribbon path for an `.xlam` at all** - not
`GetCustomUI` on ThisWorkbook, not the customUI package part - while it DOES
query a registered COM add-in for `IRibbonExtensibility`. So the ribbon is
now delivered by a small COM add-in that reads the very same `customUI14.xml`
and forwards every callback back into the VBA add-in via `Application.Run`.

### The decisive evidence
- **Every VBA-side shape was falsified** - unlocked, fully compiled
  (vbaProject.bin grew ~45 KB of p-code), locked, with/without the customUI
  part: the add-in always loads and runs (StartUp logged, hotkeys register,
  highlighting works) but `GetCustomUI` is NEVER called and no tab renders.
- **A COM add-in delivers the ribbon reliably** - implementing
  `IDTExtensibility2` + `IRibbonExtensibility` against the canonical GAC
  interop assemblies (Extensibility.dll + OFFICE.DLL; hand-rolled
  `[ComImport]` interfaces crash the CLR in-process, which is why earlier
  COM attempts failed), `GetCustomUI` is called on every launch, the full
  13.7 KB XML is delivered, `onLoad` fires and crosses back into VBA
  ("Ribbon UI attached" in the log), and the Highlighter tab appears in the
  ribbon.
- **Latent XML bugs that could never surface before** - once Excel finally
  parsed the XML it found two real schema violations that had been hiding
  since day one:
  1. `label` and `getLabel` declared together on the two exclude buttons
     ("Mutually exclusive attributes") - the dynamic getLabel wins, static
     labels removed.
  2. Six invalid `imageMso` IDs on this build (`TableSelectRow`,
     `TableSelectColumn`, `FillColor`, `DiagramColorToggle`, `ProtectSheet`,
     `CellShading`, `BorderOutsideThick`, `FlashFill`) - replaced with
     IDs that validate (a loop launched Excel, read the error dialog, and
     swapped each bad ID until the XML validated clean).
- **Defined-name syntax bug broke highlighting on Excel 2024** - creating a
  name `_XLCH_Row` throws "The syntax of this name isn't correct" (#1004) on
  modern Excel; the clean `XLCH_Row` succeeds. All defined-name constants now
  use the clean prefix, so highlighting actually works on Excel 2024 (the
  error had been silently killing every `EnsureNamesExist` call).

### Added
- **`comribbon/Connect.cs`** - the COM ribbon add-in source (single class:
  every ribbon callback implemented as a method, delegating to the VBA
  add-in via `Application.Run`; value-returning callbacks use the new
  `*Value` bridge functions in `RibbonCallbacks.bas` because `Run` passes
  arguments by value).
- **`scripts/build-comribbon.ps1`** - compiles `Connect.cs` against the
  .NET Framework csc.exe + GAC interop assemblies, deploys
  `ExcelHighlighterRibbon.dll` next to the .xlam, registers it under
  `HKCU` (RegAsm regfile rewritten to HKCU\Software\Classes - no admin
  needed), sets the Excel COM add-in keys with `LoadBehavior=3`, and
  clears any Disabled Items entry.
- **`*Value` bridge functions** in `RibbonCallbacks.bas` - thin `Function`
  wrappers over the existing ByRef `Sub` callbacks so `Application.Run` can
  return values.

### Fixed
- **`install.ps1` now builds + registers the COM ribbon add-in**
  automatically (calls `scripts/build-comribbon.ps1`), so a fresh install
  shows the Highlighter tab on Excel 2024 / recent 365 with no manual step.
- **`uninstall.ps1` removes the COM add-in registration** (ProgID, CLSID,
  Addins keys, Disabled Items).
- **`customUI/customUI14.xml` / `customUI.xml`** - removed the
  `label`+`getLabel` conflicts and replaced all `imageMso` IDs with ones
  that validate on Excel 2024.
- **`src/Constants.bas`** - defined names now use the clean `XLCH_` prefix
  (no leading underscore) so highlighting works on modern Excel; comments in
  `HighlightEngine.bas` and `Utilities.bas` updated to match.

## [2.1.3] - 2026-08-02

**The 2.1.2 "signature is the gate" conclusion was WRONG - overturned by a
decisive experiment, and the real root cause found.** The ribbon renders on
Excel 2024 (build 16.0.20228) only for add-ins whose VBA project is
PASSWORD-LOCKED (precompiled) - exactly the shape of the one add-in that
works (ASAP Utilities).

### The decisive experiment (ASAP-strip test)
Took ASAP Utilities' own file (the only add-in that renders on this
machine), stripped all three of its signature parts, and registered the
now-UNSIGNED copy in its place. Its tab still rendered, from the very same
folder our add-in lives in. **An unsigned add-in absolutely renders here -
the signature is not the gate.** (Also verified: ASAP is not a COM add-in,
and its .xlam contains no customUI part at all - it delivers its ribbon
purely via `IRibbonExtensibility.GetCustomUI`.)

### What the new experiments proved
- **The customUI package part is ignored on this build** - a minimal static
  part with zero callbacks (spec-perfect wiring: part + content-type +
  root/workbook relationships) fails to render, both in our package AND
  when injected into stripped-ASAP's package (whose own GetCustomUI ribbon
  still rendered, proving the part was ignored rather than "processed but
  empty"). `onLoad` never fires for the full XML either.
- **`GetCustomUI` is never called for an UNLOCKED project** - a minimal
  add-in containing nothing but `Implements IRibbonExtensibility` with the
  correct `Function GetCustomUI(...) As String` signature (built natively
  by Excel, same tooling as our build) had `Workbook_Open` fire (macro
  markers written via plain VBA file I/O, immune to logging failures) but
  `GetCustomUI` was NEVER called. Same result for the full 13-component
  add-in with and without the part.
- **The remaining structural difference from the working add-in is the
  locked project** - ASAP's VBA project is password-locked; ours is not.
  A locked VBA project is stored precompiled, and Excel's ribbon engine
  discovers `IRibbonExtensibility` from that compiled type info. An
  unlocked (source-only) project is apparently queried at a point where
  lazy compilation has not yet exposed the interface, so Excel silently
  treats the add-in as having no ribbon - matching every observation: no
  tab, no `onLoad`, no `GetCustomUI`, no error.

### Added
- **`scripts/lock-vba-project.ps1`** - locks the VBA project (Tools >
  VBAProject Properties > Protection > "Lock project for viewing" +
  password - the one step the object model cannot automate), then
  verifies `VBProject.Protection` from a fresh open and re-injects the
  customUI ribbon wiring that Excel's re-save strips. The ribbon then
  comes from `GetCustomUI` reading `customUI14.xml` from the add-in folder
  - exactly ASAP's delivery mechanism.

### Fixed / corrected
- **README and 2.1.2 changelog corrected** - the "signature is the gate"
  conclusion is replaced with the locked-project finding. Signing is still
  useful for the Publisher column, but it is NOT required for the ribbon.

## [2.1.2] - 2026-08-02

Investigation round that hardened the signing flow. Its "the ribbon gate is
 the VBA digital signature" conclusion was subsequently OVERTURNED by the
2.1.3 ASAP-strip experiment - see 2.1.3. The signing-flow fixes below
(clean-slate strip, fresh-open verification) remain valid and useful.

### Established (in VALID interactive sessions - real `excel.exe` launch,
probe workbook open, UIA ribbon sanity tabs Home/Insert/View + the ASAP
Utilities control tab all present)
- **The fresh unsigned build LOADS fine** - `StartUp` is logged, hotkeys
  register, macros run. The `/R` fix in 2.1.0 is confirmed working; the
  add-in itself is healthy and highlights work.
- **The ribbon is suppressed for the UNSIGNED add-in, regardless of
  delivery mechanism** - with the declarative customUI part: no tab, no
  onLoad, no GetCustomUI. With a NO_CUSTOMUI_PART=1 build (pure
  `IRibbonExtensibility.GetCustomUI`, exactly ASAP's delivery): `StartUp`
  logs but **GetCustomUI is never even called** - Excel suppresses the
  ribbon at the trust check, before any delivery mechanism is consulted.
- **ASAP Utilities (no customUI part, valid EV signature in the user's
  Trusted Publishers) renders its tab in the SAME sessions** where ours
  does not - a clean positive control.
- **The trusted-location path is NOT the gate** - the add-in's folder is
  covered by a registered Trusted Location (Location7) and macros run, yet
  the ribbon still does not render. VBAWarnings=1 runs macros everywhere,
  so macro execution never proved the location was effective; direct
  registry verification now confirms the location IS effective, and the
  ribbon still does not render.
- Conclusion: **Excel 2024 build 16.0.20228 suppresses ALL customUI
  delivery for an add-in without a valid Trusted Publisher signature.**
  The signature gate was never tested with a VALID signature in any prior
  round - every attempt landed "Stale"/invalid (Excel's own
  `VBProject.Signed=False` on a fully readable 13-component project is
  authoritative; ASAP's `False` is a locked-project artifact).

### Fixed
- **Leftover conflicting signature parts silently break the add-in
  loader** - repeated manual VBE signing with an incomplete "Remove" step
  left THREE signature parts in one file (`vbaProjectSignature.bin` +
  Agile + V3). Excel's ADD-IN LOADER refuses such a file outright (no
  StartUp in the log) while a direct workbook open tolerates it - which
  explained why, after the 13:58 signing run, the add-in stopped
  auto-loading entirely. New `Remove-XlamSignatureParts` in
  `scripts/xlam-ribbon-utils.ps1` strips every signature part (and its
  `[Content_Types].xml` Override) while preserving everything else
  byte-for-byte; `sign-xlam.ps1` now runs it before every signing attempt
  so each attempt starts from a provably clean slate.
- **`sign-xlam.ps1` verification now uses Excel's own verdict** - the
  cryptographic digest verifier is known UNRELIABLE for VBA Agile/V3
  signatures (it reported "Stale" even for ASAP Utilities, which renders
  fine - those signatures do not embed a plain MD5/SHA digest of
  `vbaProject.bin`). The script now opens the saved file in a fresh
  hidden Excel instance and reads `VBProject.Signed` on the fully readable
  project - the only trustworthy check. The digest verifier remains only
  as a fallback when the fresh open cannot read the project.
- **Misleading `VBProject.Signed` "false negative" warning removed** -
  the old text told the user the check is unreliable and to trust the
  broken verifier. The false-negative artifact applies only to LOCKED
  projects (like ASAP's). On our readable project, `Signed=False` is a
  REAL failure, and the script now says so and re-prompts the exact
  dialog steps (Remove > Choose > OK > OK > Ctrl+S).

## [2.1.1] - 2026-08-02

Follow-up round after 2.1.0: the user still reported the ribbon not
appearing. Direct OPC forensics on the live installed file - not another
launch-harness round - identified two concrete, verifiable defects that
were confounded with each other in every prior test:

### Fixed
- **Signing re-save strips the customUI relationship (a present-but-unwired
  part is worse than no part)** - verified on the installed `.xlam` via OPC
  forensics: the repo build ships the `ui/extensibility` relationship in
  BOTH `_rels/.rels` and `xl/_rels/workbook.xml.rels` (correct), but the
  installed, signed file had it ONLY at the package root. The manual
  signing step (`sign-xlam.ps1`, VBE > Tools > Digital Signature, then
  save) makes Excel RE-SAVE the add-in, and that re-save regenerates
  `xl/_rels/workbook.xml.rels` and drops the customUI relationship, because
  Excel does not understand the customUI part. The file still CONTAINS the
  part but no longer points at it from the workbook. Because a customUI
  part exists in the package, Excel never falls back to
  `IRibbonExtensibility.GetCustomUI` (the part masks the interface), while
  the unwired part is never processed either - reproducing exactly: no tab,
  onLoad never fired, GetCustomUI never called, no error.
  - New `scripts/xlam-ribbon-utils.ps1` with `Test-XlamRibbonWiring` and
    `Repair-XlamRibbonWiring`. The repair re-injects whatever wiring is
    missing (part, `[Content_Types].xml` Override, root + workbook
    relationships), copying every other package entry byte-for-byte so
    `xl/vbaProject.bin` and the `xl/vbaProjectSignature*.bin` parts are
    untouched and a valid signature stays valid.
  - `scripts/sign-xlam.ps1` now re-injects the ribbon wiring into the
    signed file AFTER the Excel save (which strips it) and BEFORE locking
    read-only, and re-verifies wiring before `-Deploy`.
  - `install.ps1` now validates the deployed file's wiring and repairs it
    in place if a stale copy slipped through.
  - `scripts/build-xlam.ps1` now asserts post-build that the part +
    content-type + both relationships are present, failing the build
    loudly if a future change could silently ship an unwired part.
- **Signature staleness was never actually tested against a valid
  signature** - certificate forensics on the live files: ASAP Utilities'
  signature blobs decode to a real EV chain whose signer cert is in the
  user's Trusted Publishers store. Ours decodes to the "Sandeep Khadka"
  self-signed cert (also in TrustedPublisher + Root), but Excel's own
  `VBProject.Signed` reads `False` on a fully readable project (13
  components), i.e. Excel treats our signature as invalid/stale. (ASAP's
  `Signed=False` is a password-locked-project artifact, not evidence the
  gate is off.) The earlier "signed build still fails" tests never
  exercised a genuinely valid signature, so the 2.1.0 signature gate
  conclusion stands - signing correctly, verifying the digest, and locking
  the file read-only (so Excel cannot re-save and invalidate it again) is
  still required for the ribbon on Excel 2024 / recent 365 builds.
- **`install.ps1` now warns when a reinstall clobbers a valid signature** -
  `install.bat` rebuilds the add-in unsigned, silently destroying any
  previously applied signature. The installer now detects signature parts
  on the existing installed copy and tells you to re-run
  `scripts/sign-xlam.ps1 -Deploy -Trust` afterwards.

## [2.1.0] - 2026-08-02

Deep-dive investigation into "add-in shows as installed and checked in
File > Options > Add-ins but the Highlighter ribbon tab never appears, no
error, on any PC / Excel version". Findings below were established in VALID
interactive sessions - a probe workbook actually open, and the ribbon tab
strip confirmed present via UI Automation sanity checks (Home/Insert/View).
(The previous round's harness launched bare `excel.exe`, which opens the
start screen where no ribbon exists, so every "no tab" result it produced
was invalid - that is how 2.0.0's part/VBA conclusions and 2.0.1's `/R`
reversal got onto the wrong track.)

### Fixed
- **Add-in never loads at all when registered with `/R` (the real,
  confirmed root cause of "shows installed/checked but no ribbon and no
  error" on every device)** - `install.ps1` wrote the Excel `OPEN`
  registry value as `/R "...excel-highlighter.xlam"`. Empirically
  verified: with `/R` the add-in shows as installed/checked but
  `Workbook_Open` never fires (no StartUp log line, no hotkeys, no
  ribbon, no error) - exactly the reported symptom. With the plain quoted
  path the add-in loads: StartUp logs, hotkeys register, COM reports
  `Installed=True, IsAddin=True`, identical to a working commercial add-in
  (ASAP Utilities, which is registered in the same OPEN key without `/R`
  and loads fine). 2.0.1's claim that `/R` is required and that removing
  it opens the file as a normal workbook was tested and disproved.
  Removed `/R`. This single fix resolves the original cross-PC/version
  complaint - every copy the user tried carried `/R` because the 1.4.0 fix
  was written up but never applied to the script (see 2.0.0).
- **Ribbon delivery: part restored as default; VBA interface kept as
  fallback** - the build previously shipped a VBA-only package (no
  customUI part) based on the invalid start-screen-harness conclusion that
  the declarative part is "silently ignored". Re-testing in valid sessions
  showed the part was never the problem: with a clean, correctly-wired
  part (Content_Types Override + relationship on both the package root and
  `xl/_rels/workbook.xml.rels`) `onLoad` still never fired - and, notably,
  a bare no-VBA add-in carrying only a minimal part failed to render too,
  while ASAP Utilities' signed add-in rendered in the same session.
  `scripts/build-xlam.ps1` now defaults to the classic part-injected
  package (`NO_CUSTOMUI_PART=1` opts out); `ThisWorkbook`'s
  `IRibbonExtensibility.GetCustomUI` remains as a fallback for builds that
  do not honour the part.
- **Excel 2024 suppresses the custom ribbon for unsigned add-ins (strongly
  indicated; requires one manual signing step)** - all valid tests on
  Excel 2024 build 16.0.20228 point to a trust gate rather than a delivery
  mechanism: every unsigned delivery path (part, VBA interface, bare
  part-only add-in) fails to render, while the one add-in whose ribbon
  does render is digitally signed by a certificate in the user's Trusted
  Publishers store. The VBA object model exposes no signing API, so
  signing is a one-time manual step (`scripts/sign-xlam.ps1`, VBE >
  Tools > Digital Signature > Sandeep Khadka > OK, then save).
  `install.ps1` now imports the publisher certificate into the current
  user's Trusted Publishers store automatically (non-elevated,
  idempotent) and the success banner tells you to run the signing script
  if the tab is still missing on Excel 2024 / recent 365 builds.

## [2.0.1] - 2026-08-01

Correction release. The 2.0.0 rebuild targeted the wrong root causes and
actually *reintroduced* both problems that make the ribbon tab silently
vanish on Excel 2024/365 (64-bit) with no error message:

### Fixed
- **Ribbon tab silently not loading (real cause #1 - reverted a 2.0.0
  regression)** - 2.0.0 added dynamic GDI swatch callbacks back to the
  ribbon (`getItemImage="GetGallery_ItemImage"` on all three colour
  galleries and `getImage="GetSwatchImage"` on the Custom Colour button).
  On 64-bit Office (Excel 2021/2024/365) `OleCreatePictureIndirect` throws
  `0x8000FFFF E_UNEXPECTED`, and Excel's RibbonX engine responds by
  silently aborting the entire custom tab - no tab, no error, on every
  device. The pre-2.0.0 XML used native `imageMso` icons and was
  documented as working. Removed the GDI callbacks from both
  `customUI/customUI14.xml` and `customUI/customUI.xml` and restored
  `imageMso="ColorPalette"` for the Custom Colour button. The swatch
  generator callbacks remain in `RibbonCallbacks.bas` / `ColourPicker.bas`
  (unreferenced) for a future safe re-enable.
- **Add-in not loading at startup (real cause #2 - reverted a 2.0.0
  regression)** - 2.0.0 removed the `/R` prefix from the Excel `OPEN`
  registry value written by `install.ps1`, claiming `/R` breaks digital
  signatures. That claim was wrong: `/R` opens the add-in read-only, which
  is exactly what prevents Excel from re-saving and invalidating a VBA
  signature, and it is the format Excel itself writes when you add an
  add-in via the UI. Without `/R`, Excel may open the `.xlam` as a normal
  read-write workbook instead of loading it as an add-in - no ribbon tab,
  no error. Restored `/R` in `install.ps1`.

## [2.0.0] - 2026-08-01

Rebuild pass targeting the reported symptoms: ribbon tab not appearing,
VBA errors, and broken/exclamation-mark icons. Verified every `onAction` /
`getImage` / etc. callback referenced in `customUI14.xml` has a matching,
correctly-signed procedure in `RibbonCallbacks.bas` (all matched - the
ribbon code itself was not the problem).

### Fixed
- **Add-in registered read-only (real root cause of signature/save issues,
  reintroduced)** - `install.ps1` was still writing the Excel `OPEN`
  registry value with a `/R` flag despite the 1.4.0 changelog claiming this
  was fixed - the fix was written up but never actually applied to the
  script. `/R` forces Excel to load the add-in read-only, which is the
  documented cause of the "Publisher blank" / signature-invalidation issue
  and can also interfere with the add-in persisting its own state. Removed.
- **Ribbon silently stops appearing after any crash ("Disabled Items")** -
  if the add-in ever threw an error while Excel was loading it (a bad
  build, a crash mid-save, an interrupted install - exactly the kind of
  thing "lots of errors" produces), Excel silently blacklists it in the
  `Resiliency\DisabledItems` registry key and will keep refusing to load
  it, including the ribbon tab, on every future launch with no visible
  error message. `install.ps1` and `uninstall.ps1` now clear any entry
  referencing this add-in. This is the most common real-world cause of a
  previously-working ribbon tab disappearing.
- **Duplicate/stray installer copies** - there were three divergent copies
  of the installer (`install.ps1`, `install/install.ps1`, and an
  extension-less `install` file), all with the same `/R` bug, out of sync
  with each other. Removed the two stray copies; `install.ps1` /
  `install.bat` are now the only installer.
- **Duplicate add-in file left in `%APPDATA%\Microsoft\AddIns\` root** -
  the installer copied the `.xlam` to two locations but only registered
  the subfolder copy in the registry, and `uninstall.ps1` only ever
  cleaned up the subfolder copy, leaving an orphaned duplicate behind
  after every uninstall. Now only one copy is written and removed.
- **Ribbon gallery regenerated GDI bitmaps on every repaint** - `getItemImage`
  (colour swatches) is polled by Excel far more often than the underlying
  colour actually changes (every repaint, hover, and `Invalidate` call).
  Each call built a fresh GDI bitmap + `IPicture` wrapper from scratch and
  discarded it. Over a long session this is the kind of thing that
  exhausts GDI handles and shows up as broken/exclamation-mark icons and
  general instability. `ColourPicker.CreateDynamicColourSwatch` now caches
  generated swatches by RGB value and reuses them; the cache is cleared on
  Reset Settings.

### Changed
- Version bumped to 2.0.0.

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
