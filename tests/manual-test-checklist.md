# Manual test checklist

VBA add-ins of this shape are impractical to unit-test in the usual sense
(no headless Excel object model available in CI without a licensed Windows
runner). Until that's set up, this checklist is the release gate - run it
before tagging a new version.

## Basic functionality

- [ ] Install the add-in, restart Excel, confirm the **Highlighter** tab
      appears automatically without touching any workbook-specific settings.
- [ ] Toggle **Highlight** on - moving the active cell highlights nothing
      until a mode is selected (mode defaults to Crosshair per Settings, so
      this should show immediately in a fresh profile).
- [ ] Switch between Row / Column / Crosshair and confirm only the expected
      cells are shaded as you move around.
- [ ] Click the currently active mode button again - highlighting turns off
      (mode becomes "None") without unticking the master toggle.
- [ ] Toggle **Highlight** off - all shading disappears immediately across
      every open workbook.

## Colour

- [ ] Each of the seven gallery swatches applies the expected colour.
- [ ] Custom Colour opens the Excel colour dialog, and the chosen colour is
      applied and marked as the selected gallery item afterward.
- [ ] Close and reopen Excel - the previously chosen colour and mode persist.

## Multi-workbook / multi-window

- [ ] Open two workbooks side by side - each highlights independently
      (crosshair position in workbook A doesn't affect workbook B).
- [ ] Open the same workbook in two windows (View → New Window) - moving in
      one window's sheet doesn't leave a stale highlight visible in the
      other unless it's also on that sheet.

## Cleanup / non-destructiveness

- [ ] With highlighting on, save the workbook, close it, then reopen it
      outside of Excel's "recently used" (i.e. genuinely closed and
      reopened) - confirm no conditional formatting rules referencing
      `_XLCH_Row`/`_XLCH_Col` remain (Home → Conditional Formatting →
      Manage Rules should be empty, or only show rules you added yourself).
- [ ] Untick the add-in in Excel Options → Add-ins while a workbook is open
      with highlighting active - confirm formatting is stripped immediately.
- [ ] Confirm `Interior.Color` of any cell you've explicitly filled yourself
      is unchanged after enabling/disabling the highlighter repeatedly.

## Merged cells & range bounds

- [ ] Select a merged range (e.g. A1:C3) - confirm the row highlight spans rows 1 to 3 and the column highlight spans columns A to C.
- [ ] Select a single cell after selecting a merged range - confirm the highlight shrinks back to single-cell width/height.

## Dark mode & swatches

- [ ] Select a custom colour - confirm recent custom colours display solid-colour swatches in the gallery (not blank icons).
- [ ] Toggle **Dark Mode** on in the Options group - confirm highlight colors adjust to a softer contrast suitable for dark themes.
- [ ] Toggle **Dark Mode** off - confirm colors return to standard vibrancy.

## Edge cases

- [ ] Protected sheet, "Allow formatting cells" unchecked - highlighter
      should silently skip that sheet (check the log file for a caught
      error rather than a MsgBox interrupting the user).
- [ ] Protected sheet, "Allow formatting cells" checked - highlighting
      should work normally.
- [ ] Chart sheet active - no error, highlighter simply does nothing until a
      worksheet is selected again.
- [ ] Very large worksheet (paste data out to something like row 500,000,
      column XFD) - confirm moving the active cell around the used area
      stays responsive, and scrolling to a far-away blank area then
      selecting a cell there brings up the highlight without a large
      recalculation stall.
- [ ] Workbook with an existing unrelated conditional formatting rule on a
      sheet - confirm enabling/disabling the highlighter never touches that
      pre-existing rule.

## Workbook exclusion

- [ ] Open a workbook, click "Exclude Workbook" in the Options group -
      highlighting disappears from that workbook immediately.
- [ ] Switch to another workbook - highlighting works normally there.
- [ ] Switch back to the excluded workbook - highlighting remains off.
- [ ] Click "Exclude Workbook" again (now labelled "Include Workbook") -
      highlighting returns to that workbook.
- [ ] Save the excluded workbook, close it, reopen it - the exclusion
      persists (the `_XLCH_Excluded` defined name travels with the file).
- [ ] Exclude the add-in's own workbook (if visible) - the button should
      be disabled since the add-in workbook is not eligible.

## Worksheet exclusion

- [ ] With a workbook active, click "Exclude Sheet" - highlighting
      disappears from the current sheet.
- [ ] Switch to another sheet in the same workbook - highlighting works.
- [ ] Switch back - highlighting remains off.
- [ ] Save the workbook, close it, reopen it - the sheet exclusion persists.

## Keyboard shortcuts

- [ ] Press Ctrl+Shift+H with highlighting on - highlighting turns off.
- [ ] Press Ctrl+Shift+H again - highlighting turns back on.
- [ ] Press Ctrl+Shift+Z after moving around - navigates to the previous cell.
- [ ] Press Ctrl+Shift+X - navigates forward again.
- [ ] Uninstall the add-in - all hotkeys stop working.

## Status bar

- [ ] With highlighting on, status bar shows "XLCH: Crosshair - Yellow".
- [ ] Toggle highlighting off - status bar clears.
- [ ] Exclude a workbook - status bar shows "XLCH: Excluded".
- [ ] Enable "Pulse" - status bar shows "XLCH: ... [Pulse]".
- [ ] Enable "Border" style - status bar shows "XLCH: ... [Border]".

## Highlight style

- [ ] Select "Fill" - cells get an interior colour highlight.
- [ ] Select "Border" - cells get a thick border highlight instead.
- [ ] Enable "Intersect" in Crosshair mode - the intersection cell gets a
      stronger tint on top of the row/column highlight.

## Protected sheets

- [ ] Enable "Allow Protected" in the ribbon.
- [ ] Activate a protected sheet (with AllowFormattingCells off) - the
      highlighter works by temporarily unprotecting, then re-protecting.
- [ ] Disable "Allow Protected" - the highlighter skips the protected sheet.

## Animated pulse

- [ ] Enable "Pulse" from the Style group.
- [ ] Move the active cell - the highlight briefly flashes/pulses.
- [ ] Disable "Pulse" - the flash stops immediately.

## Selection history

- [ ] The Back/Forward buttons in the History group navigate through
      recently visited cells.
- [ ] The buttons are disabled when there is no history to go back/forward to.
- [ ] History is per-session only (resets when Excel closes).

## Profiles

- [ ] Change mode/colour/style, then click "Save Profile" and enter a name.
- [ ] The new profile appears in the Profiles dropdown.
- [ ] Switch to a different profile - settings change immediately.
- [ ] Close and reopen Excel - the active profile persists.

## Reset

- [ ] Reset Settings restores Off / Crosshair / Yellow and updates the
      ribbon state immediately (no restart needed).
- [ ] Recent colours list is cleared on reset.
- [ ] All new options (protected, border, intersection, animated) return
      to their defaults.
