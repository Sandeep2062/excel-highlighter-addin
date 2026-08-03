# Excel Highlighter Add-in - Complete User Guide

Welcome to the **Excel Highlighter Add-in** user guide. This document explains every control, mode, shortcut, and option available in the add-in.

---

## 1. Quick Start

Once installed (see [installation.md](installation.md)), a new **Highlighter** tab appears on your Excel ribbon.

```
[ Controls ]     [ Appearance ]     [ Options ]        [ Style ]         [ History ]     [ Profiles ]      [ Help ]
- Highlight On   - Colour Gallery   - Exclude Wb       - Fill            - Back          - Dropdown        - About
- Row            - Custom Colour... - Exclude Sheet    - Border          - Forward       - Save Profile
- Column                            - Allow Protected  - Intersect
- Crosshair                         - Dark Mode        - Pulse
                                    - Reset Settings
```

- **Enable/Disable**: Click **Highlight: On / Off**, press `Ctrl+Shift+H`, or
  **right-click any cell and choose Toggle Highlighter** from the context
  menu - all three do the same thing. By default the toggle only affects
  the **current workbook** (per-workbook scope); switch the **All
  Workbooks** toggle in the Options group to apply it to every open
  workbook at once.
- **Select Mode**: Choose **Row**, **Column**, **Crosshair** (Row + Column), or **Cell**.
- **Move Active Cell**: Move your cursor or press arrow keys — the highlight follows your active cell or merged range automatically.

---

## 2. Highlight Modes

The **Controls** group provides four mode options:

| Mode | Ribbon Button | Description | Hotkey |
|---|---|---|---|
| **Off** | Highlight Toggle | Disables highlighting entirely without clearing preferences. | `Ctrl+Shift+H` |
| **Row** | Row | Highlights the entire active row across the visible/used sheet area. | Ribbon |
| **Column** | Column | Highlights the entire active column across the visible/used sheet area. | Ribbon |
| **Crosshair** | Crosshair | Highlights both the row and column simultaneously. | Ribbon |
| **Cell** | Cell | Highlights only the active cell or merged range without row/column lines. | Ribbon |

> [!TIP]
> Clicking the currently active mode button again turns highlighting off (`None` mode) without disabling the master toggle.

**Default state:** a fresh install or **Reset Settings** starts with the
highlight **off**, in **Row** mode, coloured **true yellow** `RGB(255,255,0)`
(fill style), per-workbook scope.

---

## 2.5 Scope: This Workbook vs All Workbooks

Decides what the **Highlight: On / Off** toggle affects:

- **This Workbook** (default): toggling the highlight on only turns it on in
  the workbook you are currently in. Other open workbooks stay untouched;
  turning it off in one workbook does not affect the others. Ideal when you
  work across several files and only want the crosshair in one of them.
- **All Workbooks**: toggling on turns the highlight on for *every* open
  workbook, and turning it off from any one workbook turns it off for all of
  them.

The **All Workbooks** toggle lives in the Options group. The choice is
remembered between sessions.

---

## 3. Merged Cells & Multi-Cell Selections

In **v1.3.0**, the highlighter automatically detects merged cells (`MergeArea`):

- **Merged Ranges**: When selecting a merged cell (e.g. `A1:C3`), the row highlight spans rows 1 through 3 and the column highlight spans columns A through C.
- **Dynamic Resizing**: When moving back to a single cell, the crosshair automatically contracts to 1 cell width/height.

---

## 4. Appearance & Colour Customization

The **Appearance** group lets you tailor the visual presentation:

- **Preset Colour Gallery**: Choose from 7 curated presets: **Yellow**, **Green**, **Orange**, **Cyan**, **Blue**, **Pink**, or **Grey**. Every gallery item shows a **real colour swatch** so you can see the exact tint before choosing it.
- **Custom Colour Picker**: Click **Custom Colour...** to open the native Windows colour dialog. Any RGB colour can be selected. The button itself shows the current highlight colour as a swatch.
- **Recent Colours**: Up to 4 recently picked custom colours are stored in the registry and shown as swatches at the top of the gallery.
- **Colour correctness**: prior to 2.2.0 the preset palette was stored in the wrong byte order, which is why "Yellow" rendered cyan and "Blue" rendered violet. The palette values are now correct (Yellow = `RGB(255,255,0)`, Blue = `RGB(153,204,255)`, Orange = `RGB(255,153,0)`, Cyan = `RGB(204,255,255)`, Green = `RGB(146,208,80)`, Pink = `RGB(255,204,204)`, Grey = `RGB(211,211,211)`), and both the Dark Mode tint and the Pulse animation no longer swap the red and blue channels.

---

## 5. Style & Visual Effects

The **Style** group provides advanced visual styling:

- **Fill vs. Border**:
  - **Fill**: Applies interior cell shading (default).
  - **Border**: Applies a thick cell outline instead of fill, ideal when reading dense text or numbers without obscuring background cell formatting. Works in **every** mode - Row, Column, Crosshair and Cell (older builds only drew the border in Row mode).
- **Intersect** (Crosshair mode): the *intersection cell* is the cell where the highlighted row and column cross - the cell your cursor is in. With **Intersect** enabled, that cell gets an extra accent so it stands out from the rest of the crosshair. The accent is a **separate rule with StopIfTrue**, so it always wins at the cursor cell (older builds let the row/column rules paint over it - that's why Intersect looked like it did nothing), and it is deliberately **darker** than the row/column colour so it's always visible. In Border style the accent also fills the cell, so your cursor position stays obvious even with a border-only highlight. With it off, the row and column are tinted uniformly. In **Cell** mode the intersect accent *is* the highlight.
- **Animated Pulse**: When enabled, moving the active cell triggers a subtle 3-step pulse animation (`Application.OnTime`) that flashes the highlight at the new position and restores it. Older builds scheduled the animation with a rounded timestamp that often never fired; it now uses full-precision timing and stale chains self-cancel, so the pulse reliably plays every time.

---

## 6. Options & Theme Settings

The **Options** group controls sheet/workbook scope and environment integration:

- **Dark Mode**: Toggle **Dark Mode** to adjust highlight colors for comfortable, non-glare contrast when using Excel with dark Office themes.
- **Allow Protected**: Enables highlighting on protected sheets by temporarily unprotecting them to apply conditional formatting rules and re-protecting immediately after.
- **Reset Settings**: Restores all preferences to factory defaults (Highlight: Off, Row mode, true-yellow fill, per-workbook scope).

---

## 7. Workbook & Sheet Exclusion

If you want to skip highlighting on specific files or worksheets:

- **Exclude Workbook**: Click **Exclude Workbook**. A hidden defined name (`XLCH_Excluded = 1`) is saved inside the workbook so the file stays excluded even if opened on another machine.
- **Exclude Sheet**: Click **Exclude Sheet**. A worksheet-scoped defined name (`XLCH_SheetExcluded = 1`) excludes only the active worksheet while allowing highlighting on other sheets in the same workbook.

---

## 7.5 Right-Click Context Menu

You don't need the ribbon (or even the hotkey) to switch the highlighter on
and off. **Right-click any cell** and you'll find a **Toggle Highlighter**
item at the bottom of Excel's standard context menu:

- Click it to toggle the highlight exactly like the **Highlight: On / Off**
  ribbon button or the `Ctrl+Shift+H` hotkey.
- It shows a **check mark** when the highlight is on for the current
  workbook, so you can see the state at a glance before clicking.
- It respects the same **scope** setting: per-workbook by default, or all
  workbooks when **All Workbooks** is enabled.
- Toggling from the context menu (or the hotkey) **immediately updates the
  ribbon** - the Highlight button's pressed state and "On/Off" label follow
  the workbook you toggled, no restart needed.
- It's added automatically when the add-in loads and removed when Excel
  closes, so it never leaves anything behind.

> [!NOTE]
> In the Visual Basic Editor the add-in's project is named **Highlighter**
> (shown as `Highlighter (excel-highlighter.xlam)` in the Project Explorer,
> not the default `VBAProject`). This is a cosmetic project name only - no
> code anywhere references it.

## 8. Selection History Navigation

The **History** group tracks your cell navigation path across sheets and workbooks:

| Action | Ribbon Button | Hotkey |
|---|---|---|
| **Go Back** | Back | `Ctrl+Shift+Z` |
| **Go Forward** | Forward | `Ctrl+Shift+X` |

- Remembers up to 20 recently selected cells per session.
- Allows jump navigation across sheets and open workbooks seamlessly.

---

## 9. Profiles

The **Profiles** group lets you save and switch named configuration bundles:

1. Configure your desired mode, colour, style, intersection, and theme settings.
2. Click **Save Profile**, enter a name (e.g. *Data Entry*, *Night Audit*).
3. Switch profiles instantly at any time from the **Profiles** dropdown menu.

---

## 10. Frequently Asked Questions (FAQ)

**Q: Does this add-in alter my actual cell formatting or colors?**  
*A:* **No.** Highlighting uses non-destructive conditional formatting overlays referencing temporary defined names. When you close the workbook or unload the add-in, all formatting is cleanly removed.

**Q: Will saving the workbook save the crosshair?**  
*A:* Under normal closing operations, conditional formatting and defined names are cleaned up on `WorkbookBeforeClose`. If Excel crashes or process is killed, any residual CF rules can be cleared by turning the add-in off and on or running Reset.

**Q: Is 64-bit Excel supported?**  
*A:* **Yes.** Version 1.3.0 includes full 64-bit API declarations (`PtrSafe` / `LongPtr`) for Windows Common Dialogs and GDI swatch rendering.
