# Architecture

## Overview

The add-in has four moving pieces:

1. **`EventApp`** - a class module with `Public WithEvents App As Application`,
   kept alive for the whole session by `AddinHost.gEventApp`. It's the only
   thing that hears about selection changes, workbook opens, etc.
2. **`HighlightEngine`** - translates those events into conditional
   formatting changes. This is where almost all the interesting logic lives.
3. **`Settings`** - persisted user preferences, backed by
   `SaveSetting`/`GetSetting` (registry, per-user).
4. **`RibbonCallbacks`** - the UI layer. Reads/writes `Settings`, tells
   `HighlightEngine` to reapply, and asks the ribbon to invalidate itself.

## Why conditional formatting instead of `Interior.Color`

Setting `Interior.Color` directly is destructive: to remove the highlight
later you'd need to remember and restore whatever fill was there before,
including "no fill", themed fills, and fills coming from a table style. Any
gap in that bookkeeping (a crash, `Application.EnableEvents = False` left on
by another add-in mid-operation, etc.) leaves a permanently discoloured cell
behind.

Conditional formatting sidesteps the problem entirely: it's a layer on top of
the cell's real formatting, it's trivial to identify and remove later (we tag
our rules by referencing our own defined names in the formula), and Excel
already knows how to redraw it efficiently as part of normal screen updates.

## Why defined names instead of recalculating formulas directly

An earlier prototype used `CELL("row")` / `CELL("col")` directly in the CF
formula and forced a recalculation on every `SheetSelectionChange`. It works,
but `CELL()` without arguments is workbook-volatile - forcing a calculation
to refresh it on every single selection change means paying for a full
dependency-tree recalculation pass on every arrow-key press, which is
noticeably slow on workbooks with any non-trivial formula load.

Four hidden, workbook-scoped defined names (`_XLCH_Row`, `_XLCH_RowEnd`, `_XLCH_Col`, `_XLCH_ColEnd`)
store the start and end row/column boundaries of the active selection. For single-cell selections,
`_XLCH_Row` equals `_XLCH_RowEnd` and `_XLCH_Col` equals `_XLCH_ColEnd`. For merged cell ranges,
the names automatically capture the full merged dimensions (`MergeArea`).
Repointing these defined names' `RefersToR1C1` to new constants on `SheetSelectionChange`
only dirties the set of cells depending on those names - our CF rules - keeping the hot path cheap.

## Why the CF target range is bounded

`FormatConditions.Add` is applied to a specific `Range`. Nothing stops you
from applying it to `ws.Cells` (the entire sheet, ~17 billion cells), and in
practice Excel's renderer is usually fine with that because it only actually
evaluates CF for cells it's about to draw. But "usually fine" isn't a
guarantee across all Excel builds and hardware, and defining rules against
the full grid also means every future `UsedRange`-altering operation on the
sheet (paste, autofit, etc.) has more surface area to interact with.

Instead, `HighlightEngine.BoundedTargetRange` unions:

- the sheet's `UsedRange` (so highlighting works anywhere you already have
  data), with
- the active window's `VisibleRange` (so highlighting always works wherever
  you're currently looking, even if you've scrolled past the used range into
  blank space).

This is rebuilt whenever CF rules are rebuilt (first visit to a sheet, or a
mode/colour/enabled change) - not on every selection change, so it doesn't
sit on the hot path.

**Known trade-off:** if you scroll to an empty area that's both outside the
used range and wasn't part of the visible range at the time CF was last
rebuilt, the highlight won't show there until that area is visited again
(which retriggers a rebuild via `SheetSelectionChange`). In practice this is
rarely noticeable because moving there at all fires the event that fixes it.

## Application events used

| Event | Why |
|---|---|
| `SheetSelectionChange` | The core trigger - moves the crosshair. |
| `WorkbookOpen` | Pre-creates the two defined names so no `#NAME?` flash is possible. |
| `WorkbookActivate` | Same, for workbooks that were already open when Excel gained focus. |
| `WindowActivate` | Covers the case of two windows onto the same workbook. |
| `SheetDeactivate` | Reserved hook - currently a no-op (see future-features.md). |
| `WorkbookBeforeClose` | Cleanup: removes our CF rules and defined names before the file can be saved/closed with them present. |

## Settings storage

`SaveSetting`/`GetSetting` write to
`HKEY_CURRENT_USER\Software\VB and VBA Program Settings\ExcelCrosshairHighlighter\General`.
This was chosen over document properties or workbook names because
preferences here are meant to be **global to the user**, not tied to any one
workbook - you shouldn't have to reconfigure your preferred colour in every
file you open.

## Memory management / cleanup

- `AddinHost.gEventApp` is the one long-lived object reference. Losing it
  (e.g. assigning `Nothing` without a replacement) silently stops every
  event from firing - this is the classic bug in this style of add-in, which
  is why it's centralised in one module with a comment calling it out.
- `HighlightEngine`'s tracker dictionary is keyed by `Workbook.Name` +
  `Worksheet.CodeName` and is process-lifetime only; it's not persisted and
  doesn't need to be - it just avoids rebuilding CF on every selection
  change for sheets that are already correctly configured.
- `WorkbookBeforeClose` and `Workbook_AddinUninstall` both call the same
  cleanup path, so there's exactly one place that knows how to fully undo
  everything this add-in does to a workbook.
