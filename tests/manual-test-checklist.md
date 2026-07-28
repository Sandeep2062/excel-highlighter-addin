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

## Reset

- [ ] Reset Settings restores Off / Crosshair / Yellow and updates the
      ribbon state immediately (no restart needed).
