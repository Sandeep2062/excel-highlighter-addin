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

- **Enable/Disable**: Click **Highlight: On / Off** (or press `Ctrl+Shift+H`) to toggle highlighting across all open workbooks.
- **Select Mode**: Choose **Row**, **Column**, or **Crosshair** (Row + Column).
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

---

## 3. Merged Cells & Multi-Cell Selections

In **v1.3.0**, the highlighter automatically detects merged cells (`MergeArea`):

- **Merged Ranges**: When selecting a merged cell (e.g. `A1:C3`), the row highlight spans rows 1 through 3 and the column highlight spans columns A through C.
- **Dynamic Resizing**: When moving back to a single cell, the crosshair automatically contracts to 1 cell width/height.

---

## 4. Appearance & Colour Customization

The **Appearance** group lets you tailor the visual presentation:

- **Preset Colour Gallery**: Choose from 7 curated presets: **Yellow**, **Green**, **Orange**, **Cyan**, **Blue**, **Pink**, or **Grey**.
- **Custom Colour Picker**: Click **Custom Colour...** to open the native Windows colour dialog. Any RGB colour can be selected.
- **Recent Colours**: Up to 4 recently picked custom colours are stored in registry and rendered with dynamic solid-colour GDI swatches in the gallery dropdown.

---

## 5. Style & Visual Effects

The **Style** group provides advanced visual styling:

- **Fill vs. Border**:
  - **Fill**: Applies interior cell shading (default).
  - **Border**: Applies thick cell outlines instead of fill, ideal when reading dense text or numbers without obscuring background cell formatting.
- **Intersection Accent**: When enabled in Crosshair mode, the cell at the row/column crossing receives an extra accent tint for emphasis.
- **Animated Pulse**: When enabled, moving the active cell triggers a subtle 3-step pulse animation (`Application.OnTime`) to highlight the new position.

---

## 6. Options & Theme Settings

The **Options** group controls sheet/workbook scope and environment integration:

- **Dark Mode**: Toggle **Dark Mode** to adjust highlight colors for comfortable, non-glare contrast when using Excel with dark Office themes.
- **Allow Protected**: Enables highlighting on protected sheets by temporarily unprotecting them to apply conditional formatting rules and re-protecting immediately after.
- **Reset Settings**: Restores all preferences to factory defaults (Highlight: Off, Crosshair mode, Yellow fill).

---

## 7. Workbook & Sheet Exclusion

If you want to skip highlighting on specific files or worksheets:

- **Exclude Workbook**: Click **Exclude Workbook**. A hidden defined name (`_XLCH_Excluded = 1`) is saved inside the workbook so the file stays excluded even if opened on another machine.
- **Exclude Sheet**: Click **Exclude Sheet**. A worksheet-scoped defined name (`_XLCH_SheetExcluded = 1`) excludes only the active worksheet while allowing highlighting on other sheets in the same workbook.

---

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
