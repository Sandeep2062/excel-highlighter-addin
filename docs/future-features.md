# Future features / roadmap

Rough ideas, not commitments, roughly ordered by how likely they are to
actually be worth doing next.

## Fairly likely

- **Workbook exclusion list** - a per-workbook "don't highlight this one"
  toggle, stored via a workbook-scoped defined name so it travels with the
  file itself rather than living in the global registry settings. Referenced
  as a TODO in `HighlightEngine`.
- **Worksheet exclusion list** - same idea, one level down, for workbooks
  where only a couple of sheets are dashboards/print layouts that shouldn't
  ever show the overlay.
- **Native colour picker without the scratch-palette-slot workaround** -
  the current custom colour flow borrows palette slot 56 from the active
  workbook. It restores it afterwards, but a cleaner approach (e.g. a small
  custom UserForm with a colour wheel, or the Common Dialog API) would avoid
  touching the workbook's palette at all.
- **Keyboard shortcut to toggle on/off** - RibbonX doesn't support key
  bindings directly; this would need `Application.OnKey`, registered/
  unregistered alongside the event sink lifecycle in `AddinHost`.
- **Recent colours** - remember the last few custom RGB values chosen and
  surface them as extra gallery items.

## Plausible, needs more thought

- **Protected-sheet support** - right now we just skip sheets where
  `AllowFormattingCells` is off. Excel does allow unlocking *just* that one
  permission on a protected sheet, but no in-application handling should
  ever silently change a user's protection settings, this would need to be
  an explicit, well-labelled opt-in.
- **Merged-cell improvements** - `ROW()`/`COLUMN()` inside a merged range
  behave a bit differently depending on which cell of the merge Excel treats
  as "active"; worth a dedicated pass with real test cases rather than a
  guess.
- **Border-only highlight style** - an alternative to fill colour, for
  people who find a solid fill too visually heavy. Would need a second CF
  rule type (`Borders`) alongside the existing `Interior` one.
- **Glow / gradient highlight** - visually nicer, but conditional formatting
  doesn't support soft edges - would likely require switching the "Temporary
  display layer" approach mentioned in the original brief (floating
  transparent shapes positioned over the active row/column) for this style
  specifically, which is a meaningfully different implementation from the CF
  approach used everywhere else.
- **Status bar integration** - show the current mode/colour in the status
  bar as a quick sanity check without needing to look at the ribbon.

## Longer shots

- **Animated highlight** - a brief flash/pulse when the selection changes.
  Would need a timer loop (`Application.OnTime`) driving repeated CF colour
  changes, which fights against the "avoid unnecessary screen updates"
  performance goal, so this would need to be strictly opt-in and probably
  capped to a couple of iterations rather than a continuous animation.
- **Dark mode / theme awareness** - read `Application.OperatingSystem` /
  Office theme and adjust the default palette suggestions accordingly,
  rather than the fixed RGB set used today.
- **Multiple simultaneous highlight styles** - e.g. a light "row" tint plus
  a stronger "crosshair cell" tint at the intersection. Doable with a third
  CF rule but adds real complexity to `RebuildConditionalFormatting` for a
  fairly niche visual preference.
- **Profiles** - named bundles of mode+colour+exclusions that can be swapped
  quickly (e.g. "review mode" vs "data entry mode"). Would sit on top of the
  exclusion-list work above rather than being independent of it.
- **Selection history** - a small backward/forward navigation of recently
  visited cells, unrelated to highlighting itself but a natural fit for the
  same Application-event infrastructure.

## Explicitly not planned

- Anything that would require permanently modifying `Interior.Color` or
  other real cell formatting. That's a hard architectural boundary for this
  project, not just a current limitation.
