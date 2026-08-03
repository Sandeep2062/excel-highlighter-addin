Attribute VB_Name = "HighlightEngine"
'===============================================================================
' Module    : HighlightEngine
' Purpose   : Everything that actually paints (and un-paints) the crosshair.
'
'             Design notes
'             ------------
'             We never touch Interior.Color directly - that would permanently
'             clobber whatever formatting the user already has. Instead each
'             monitored workbook gets two hidden, workbook-scoped defined
'             names (XLCH_Row / XLCH_Col) that hold the active row/column as
'             plain numeric constants. Conditional formatting rules on the
'             worksheet reference those names (e.g. "=ROW()=XLCH_Row"), so
'             on every selection change we only need to update two Name
'             values - a very cheap operation - rather than add/remove CF
'             rules on every keystroke. CF rules themselves are only
'             (re)built when a sheet is seen for the first time, or when the
'             user changes mode/colour/enabled state from the ribbon.
'
'             NOTE: the names use the clean "XLCH_" prefix (no leading
'             underscore) - modern Excel 2024 defined-name parsers reject
'             leading-underscore names like _XLCH_Row with error 1004. See
'             Constants.bas for the full story.
'
'             To keep this workable on very large sheets (1,048,576 rows x
'             16,384 columns) the CF rules are scoped to the union of the
'             sheet's UsedRange and the current window's VisibleRange, rather
'             than the entire sheet. See docs/architecture.md for the
'             reasoning and the known limitation this introduces.
'
'             v1.1.0 additions:
'             - Border-only highlight style (alternative to fill)
'             - Intersection accent cell (stronger tint at row/col crossing)
'             - Protected-sheet support (opt-in, unprotects temporarily)
'             - Animated pulse on selection change
'             - Merged-cell awareness
'===============================================================================
Option Explicit

' Tracks which worksheets currently have our CF applied, and with what
' mode+colour signature, so we don't rebuild CF on every SelectionChange.
' Key   : Workbook.Name & "|" & Worksheet.Name
' Value : "<mode>|<rgb>" signature string
Private mTrackers As Object

' Session-only per-workbook "highlight on" markers. Used only when
' Settings.ScopeAll is False (the default): the toggle then affects the
' active workbook only, and this dictionary remembers which workbooks have
' been switched on, so turning the highlight off in workbook B does not
' switch it off in workbook A. Key: wb.Name (unique within an instance).
Private mWorkbookEnabled As Object

' Pulse "generation" counter. Every TriggerAnimationPulse bumps it; each
' scheduled PulseStep carries the generation it was launched with and
' immediately exits if a newer pulse has started. This cleanly kills stale
' timer chains without relying on Application.OnTime cancellation (which
' requires the exact same EarliestTime + procedure string to be passed back,
' and silently no-ops if the string differs by even one character).
Private mPulseGeneration As Long

' Separator used inside the pulse "workbook/worksheet" payload passed through
' Application.OnTime's procedure string. NOT the pipe character - sheet names
' may legally contain "|", which would corrupt Split(). vbVerticalTab (Chr 11)
' is a built-in constant (legal in a Const - a Chr$() call would be a "constant
' expression required" compile error) and is not typeable in either a sheet
' tab name or a file name.
Private Const PULSE_SEP As String = vbVerticalTab

'-------------------------------------------------------------------------------
' EnsureTrackers
'-------------------------------------------------------------------------------
Private Sub EnsureTrackers()
    If mTrackers Is Nothing Then
        Set mTrackers = CreateObject("Scripting.Dictionary")
    End If
    If mWorkbookEnabled Is Nothing Then
        Set mWorkbookEnabled = CreateObject("Scripting.Dictionary")
    End If
End Sub

'-------------------------------------------------------------------------------
' IsWorkbookHighlightActive
' The single gate that decides whether a workbook should be highlighted right
' now. With ScopeAll=True ("apply to all open workbooks") any eligible
' workbook is active whenever Settings.enabled is on. With ScopeAll=False
' (the default, per-workbook), only workbooks the user has switched on in
' this session are active - other open workbooks stay untouched.
'-------------------------------------------------------------------------------
Public Function IsWorkbookHighlightActive(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    If Not Settings.enabled Then Exit Function
    If Settings.ScopeAll Then
        IsWorkbookHighlightActive = True
    Else
        IsWorkbookHighlightActive = mWorkbookEnabled.Exists(wb.name)
    End If
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' SetWorkbookEnabled / WorkbookEnabled / ClearWorkbookEnabled
' Per-workbook toggle state (only meaningful when ScopeAll is False).
'-------------------------------------------------------------------------------
Public Sub SetWorkbookEnabled(ByVal wb As Workbook, ByVal onState As Boolean)
    EnsureTrackers
    If onState Then
        mWorkbookEnabled(wb.name) = True
    Else
        If mWorkbookEnabled.Exists(wb.name) Then mWorkbookEnabled.Remove wb.name
    End If
End Sub

Public Function WorkbookEnabled(ByVal wb As Workbook) As Boolean
    EnsureTrackers
    On Error Resume Next
    WorkbookEnabled = mWorkbookEnabled.Exists(wb.name)
    On Error GoTo 0
End Function

Public Sub ClearWorkbookEnabled(ByVal wb As Workbook)
    EnsureTrackers
    If mWorkbookEnabled.Exists(wb.name) Then mWorkbookEnabled.Remove wb.name
End Sub

'-------------------------------------------------------------------------------
' ActiveWorkbookHighlightState
' Convenience for the ribbon getPressed/getLabel callbacks: the effective
' on/off state of the workbook that currently owns the ribbon (ActiveWorkbook
' when the callback runs).
'-------------------------------------------------------------------------------
Public Function ActiveWorkbookHighlightState() As Boolean
    On Error Resume Next
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Function
    ActiveWorkbookHighlightState = IsWorkbookHighlightActive(wb)
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' HandleSelectionChange
' Description : Main entry point, called by EventApp on every
'               Application.SheetSelectionChange.
' Parameters  : sh     - the sheet reported by the event (may be a chart sheet)
'               target - the new selection
'-------------------------------------------------------------------------------
Public Sub HandleSelectionChange(ByVal sh As Object, ByVal target As Range)

    On Error GoTo ErrHandler

    If Not Settings.enabled Then
        UpdateStatusBar Nothing
        Exit Sub
    End If
    If Settings.Mode = hmNone Then
        UpdateStatusBar Nothing
        Exit Sub
    End If
    If Not Utilities.SheetIsEligible(sh) Then Exit Sub

    Dim ws As Worksheet
    Set ws = sh

    Dim wb As Workbook
    Set wb = ws.Parent

    If Not Utilities.WorkbookIsEligible(wb) Then Exit Sub

    ' Per-workbook scope: skip workbooks that were never toggled on.
    If Not IsWorkbookHighlightActive(wb) Then Exit Sub

    ' Check per-workbook exclusion before doing anything else.
    If Utilities.WorkbookIsExcluded(wb) Then
        UpdateStatusBar wb
        Exit Sub
    End If

    ' Check per-sheet exclusion.
    If Utilities.SheetIsExcluded(ws) Then
        UpdateStatusBar wb
        Exit Sub
    End If

    ' Record the selection in history for back/forward navigation.
    SelectionHistory.Push wb.name, ws.name, target.Row, target.Column

    EnsureNamesExist wb

    ' Cheap part: just move the crosshair. Runs on every single selection
    ' change, so it must stay fast even on huge sheets.
    UpdatePositionNames wb, target

    ' Expensive part: only runs when this sheet hasn't been configured yet
    ' for the current mode/colour (first visit, or settings just changed).
    If Not SheetMatchesCurrentSignature(wb, ws) Then
        RebuildConditionalFormatting ws
    End If

    ' Update the status bar with current state.
    UpdateStatusBar wb

    ' Trigger animation pulse if enabled.
    If Settings.AnimatedEnabled Then
        TriggerAnimationPulse ws
    End If

    Exit Sub

ErrHandler:
    Logging.LogError "HighlightEngine.HandleSelectionChange", Err.Number, Err.Description, sh.name

End Sub

'-------------------------------------------------------------------------------
' HandleWorkbookOpen / HandleWorkbookActivate / HandleWindowActivate
' Description : These don't need to do much - the real work happens lazily on
'               the first SheetSelectionChange - but we proactively make sure
'               names exist so a freshly opened workbook doesn't show a
'               #NAME? flash if a CF rule somehow evaluates before the first
'               selection event fires (e.g. opening straight into a sheet
'               that already had a selection restored from the last save).
'-------------------------------------------------------------------------------
Public Sub HandleWorkbookOpen(ByVal wb As Workbook)
    On Error GoTo ErrHandler
    If Not Utilities.WorkbookIsEligible(wb) Then Exit Sub
    If Not IsWorkbookHighlightActive(wb) Then Exit Sub
    If Utilities.WorkbookIsExcluded(wb) Then Exit Sub
    EnsureNamesExist wb
    Exit Sub
ErrHandler:
    Logging.LogError "HighlightEngine.HandleWorkbookOpen", Err.Number, Err.Description, wb.name
End Sub

Public Sub HandleWorkbookActivate(ByVal wb As Workbook)
    On Error GoTo ErrHandler
    If Not Utilities.WorkbookIsEligible(wb) Then Exit Sub
    If Not IsWorkbookHighlightActive(wb) Then Exit Sub
    If Utilities.WorkbookIsExcluded(wb) Then Exit Sub
    EnsureNamesExist wb
    Exit Sub
ErrHandler:
    Logging.LogError "HighlightEngine.HandleWorkbookActivate", Err.Number, Err.Description, wb.name
End Sub

Public Sub HandleWindowActivate(ByVal wb As Workbook)
    ' Same shape as WorkbookActivate - kept as a separate entry point so the
    ' two event sources stay independently traceable in the log.
    On Error GoTo ErrHandler
    If Not Utilities.WorkbookIsEligible(wb) Then Exit Sub
    If Not IsWorkbookHighlightActive(wb) Then Exit Sub
    If Utilities.WorkbookIsExcluded(wb) Then Exit Sub
    EnsureNamesExist wb
    Exit Sub
ErrHandler:
    Logging.LogError "HighlightEngine.HandleWindowActivate", Err.Number, Err.Description, wb.name
End Sub

'-------------------------------------------------------------------------------
' HandleSheetDeactivate
' Description : No repaint needed - CF formulas already only evaluate true for
'               one row/column, and the workbook-level name is shared across
'               sheets, so nothing needs to be un-drawn. Reserved as a hook
'               point for a future per-sheet "clear on leave" preference.
'-------------------------------------------------------------------------------
Public Sub HandleSheetDeactivate(ByVal sh As Object)
    ' NOTE: intentionally a no-op today. See docs/future-features.md.
End Sub

'-------------------------------------------------------------------------------
' HandleWorkbookBeforeClose
' Description : Removes everything we added to a workbook before it closes -
'               the add-in must never leave a workbook permanently modified.
'               This is what makes "never permanently modify formatting" true
'               even if the user saves right before closing.
'-------------------------------------------------------------------------------
Public Sub HandleWorkbookBeforeClose(ByVal wb As Workbook)

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        RemoveOurConditionalFormatting ws
        UntrackSheet wb, ws
    Next ws

    Utilities.SafeDeleteName wb, NAME_ROW_PREFIX
    Utilities.SafeDeleteName wb, NAME_ROW_END_PREFIX
    Utilities.SafeDeleteName wb, NAME_COL_PREFIX
    Utilities.SafeDeleteName wb, NAME_COL_END_PREFIX
    Utilities.DeleteLegacyNames wb

    ClearWorkbookEnabled wb

    Exit Sub

ErrHandler:
    ' Don't block the close on a logging or cleanup failure.
    Logging.LogError "HighlightEngine.HandleWorkbookBeforeClose", Err.Number, Err.Description, wb.name

End Sub

'-------------------------------------------------------------------------------
' ReapplyAllOpenWorkbooks
' Description : Called by RibbonCallbacks whenever the user changes mode,
'               colour or the enabled toggle. Forces every sheet that has
'               previously been configured (i.e. has a tracker entry) plus
'               each window's current active sheet to be rebuilt against the
'               new settings. Sheets that were never visited pick up the new
'               settings automatically the first time they are, via
'               HandleSelectionChange, so we don't need to touch every sheet
'               in every open workbook here.
'-------------------------------------------------------------------------------
Public Sub ReapplyAllOpenWorkbooks()

    On Error GoTo ErrHandler

    EnsureTrackers

    Dim wb As Workbook
    Dim ws As Worksheet

    For Each wb In Application.Workbooks

        If Utilities.WorkbookIsEligible(wb) Then

            If Not Settings.enabled Then
                ' Turning the add-in off: strip everything back out.
                For Each ws In wb.Worksheets
                    RemoveOurConditionalFormatting ws
                    UntrackSheet wb, ws
                Next ws
                Utilities.SafeDeleteName wb, NAME_ROW_PREFIX
                Utilities.SafeDeleteName wb, NAME_ROW_END_PREFIX
                Utilities.SafeDeleteName wb, NAME_COL_PREFIX
                Utilities.SafeDeleteName wb, NAME_COL_END_PREFIX
                Utilities.DeleteLegacyNames wb
                ' Also forget any per-workbook "on" marker - the whole
                ' highlighter is off.
                ClearWorkbookEnabled wb
            ElseIf Not IsWorkbookHighlightActive(wb) Then
                ' Per-workbook scope (ScopeAll=False) and this workbook was
                ' never toggled on: strip any leftover highlight so a
                ' previously-enabled workbook being toggled off loses its
                ' paint, while other workbooks keep theirs.
                For Each ws In wb.Worksheets
                    RemoveOurConditionalFormatting ws
                    UntrackSheet wb, ws
                Next ws
                Utilities.SafeDeleteName wb, NAME_ROW_PREFIX
                Utilities.SafeDeleteName wb, NAME_ROW_END_PREFIX
                Utilities.SafeDeleteName wb, NAME_COL_PREFIX
                Utilities.SafeDeleteName wb, NAME_COL_END_PREFIX
                Utilities.DeleteLegacyNames wb
            Else
                EnsureNamesExist wb
                ' Rebuild only sheets we already know about, plus the sheet
                ' currently on screen for this workbook's active window.
                For Each ws In wb.Worksheets
                    If IsTracked(wb, ws) Then
                        RebuildConditionalFormatting ws
                    End If
                Next ws

                On Error Resume Next
                If TypeOf wb.ActiveSheet Is Worksheet Then
                    RebuildConditionalFormatting wb.ActiveSheet
                End If
                On Error GoTo ErrHandler
            End If

        End If

    Next wb

    Exit Sub

ErrHandler:
    Logging.LogError "HighlightEngine.ReapplyAllOpenWorkbooks", Err.Number, Err.Description

End Sub

'===============================================================================
' Internal helpers
'===============================================================================

'-------------------------------------------------------------------------------
' EnsureNamesExist
' Creates the two hidden position names if this workbook doesn't have them
' yet. Cheap no-op on subsequent calls.
'-------------------------------------------------------------------------------
Public Sub EnsureNamesExist(ByVal wb As Workbook)
' NOTE: Made Public so RibbonCallbacks.OnExclude_Action can call it when
' re-enabling highlighting on a previously excluded workbook.

    If Not Utilities.NameExists(wb, NAME_ROW_PREFIX) Then
        wb.Names.Add name:=NAME_ROW_PREFIX, RefersToR1C1:="=1", Visible:=False
    End If

    If Not Utilities.NameExists(wb, NAME_ROW_END_PREFIX) Then
        wb.Names.Add name:=NAME_ROW_END_PREFIX, RefersToR1C1:="=1", Visible:=False
    End If

    If Not Utilities.NameExists(wb, NAME_COL_PREFIX) Then
        wb.Names.Add name:=NAME_COL_PREFIX, RefersToR1C1:="=1", Visible:=False
    End If

    If Not Utilities.NameExists(wb, NAME_COL_END_PREFIX) Then
        wb.Names.Add name:=NAME_COL_END_PREFIX, RefersToR1C1:="=1", Visible:=False
    End If

End Sub

'-------------------------------------------------------------------------------
' UpdatePositionNames
' The hot path - runs on every selection change. Repoints position names.
' Detects merged ranges to ensure crosshairs cover full merged dimensions.
'-------------------------------------------------------------------------------
Private Sub UpdatePositionNames(ByVal wb As Workbook, ByVal target As Range)

    On Error GoTo Fallback

    Dim rStart As Long, rEnd As Long, cStart As Long, cEnd As Long

    If target.MergeCells Then
        Dim ma As Range
        Set ma = target.MergeArea
        rStart = ma.Row
        rEnd = ma.Row + ma.Rows.count - 1
        cStart = ma.Column
        cEnd = ma.Column + ma.Columns.count - 1
    Else
        rStart = target.Row
        rEnd = target.Row + target.Rows.count - 1
        cStart = target.Column
        cEnd = target.Column + target.Columns.count - 1
    End If

    wb.Names(NAME_ROW_PREFIX).RefersToR1C1 = "=" & rStart
    wb.Names(NAME_ROW_END_PREFIX).RefersToR1C1 = "=" & rEnd
    wb.Names(NAME_COL_PREFIX).RefersToR1C1 = "=" & cStart
    wb.Names(NAME_COL_END_PREFIX).RefersToR1C1 = "=" & cEnd
    Exit Sub

Fallback:
    On Error Resume Next
    wb.Names(NAME_ROW_PREFIX).RefersToR1C1 = "=" & target.Row
    wb.Names(NAME_ROW_END_PREFIX).RefersToR1C1 = "=" & target.Row
    wb.Names(NAME_COL_PREFIX).RefersToR1C1 = "=" & target.Column
    wb.Names(NAME_COL_END_PREFIX).RefersToR1C1 = "=" & target.Column
    On Error GoTo 0

End Sub

'-------------------------------------------------------------------------------
' RebuildConditionalFormatting
' Description : Strips any of our existing rules from the sheet and, if the
'               add-in is enabled, adds fresh rules matching the current
'               mode and colour. Bounded to UsedRange ∪ VisibleRange to keep
'               this affordable on very large worksheets.
'-------------------------------------------------------------------------------
Public Sub RebuildConditionalFormatting(ByVal ws As Worksheet)
' NOTE: Made Public so RibbonCallbacks.OnExclude_Action can rebuild the
' active sheet when a workbook is un-excluded.

    On Error GoTo ErrHandler

    ' Handle protected sheets: if AllowProtected is on, temporarily unprotect.
    Dim wasProtected As Boolean
    If ws.ProtectContents And Settings.AllowProtected Then
        On Error Resume Next
        ws.Unprotect
        wasProtected = (Err.Number = 0)
        On Error GoTo ErrHandler
    End If

    RemoveOurConditionalFormatting ws

    If Not Settings.enabled Or Settings.Mode = hmNone Then
        UntrackSheet ws.Parent, ws
        GoTo RestoreProtection
    End If

    Dim targetRange As Range
    Set targetRange = BoundedTargetRange(ws)
    If targetRange Is Nothing Then GoTo RestoreProtection

    Dim rowColour As Long, colColour As Long
    rowColour = Settings.EffectiveRowRGB
    colColour = Settings.EffectiveColRGB

    Dim style As HighlightStyle
    style = Settings.HighlightStyle

    Dim rowExpr As String, colExpr As String
    rowExpr = "AND(ROW()>=" & NAME_ROW_PREFIX & ",ROW()<=" & NAME_ROW_END_PREFIX & ")"
    colExpr = "AND(COLUMN()>=" & NAME_COL_PREFIX & ",COLUMN()<=" & NAME_COL_END_PREFIX & ")"

    Select Case Settings.Mode
        Case hmRow
            AddRule targetRange, "=" & rowExpr, rowColour, style
        Case hmColumn
            AddRule targetRange, "=" & colExpr, colColour, style
        Case hmCrosshair
            ' When per-mode colours are enabled, row and column can differ.
            AddRule targetRange, "=" & rowExpr, rowColour, style
            AddRule targetRange, "=" & colExpr, colColour, style
            ' Add intersection accent if enabled. Deliberately a SEPARATE rule
            ' from AddRule: it must StopIfTrue so it always wins at the exact
            ' cursor cell regardless of Excel's rule-precedence quirks, and in
            ' Border style it also fills the cell so it stays visible.
            If Settings.IntersectionEnabled Then
                AddIntersectionRule targetRange, "=AND(" & rowExpr & "," & colExpr & ")", _
                                    IntersectionAccentRGB(), style
            End If
        Case hmCell
            AddRule targetRange, "=AND(" & rowExpr & "," & colExpr & ")", rowColour, style
    End Select

    TrackSheet ws.Parent, ws

RestoreProtection:
    If wasProtected Then
        On Error Resume Next
        ws.Protect
        On Error GoTo ErrHandler
    End If

    Exit Sub

ErrHandler:
    ' Common cause: a protected sheet that doesn't allow formatting changes.
    Logging.LogError "HighlightEngine.RebuildConditionalFormatting", Err.Number, Err.Description, ws.name

End Sub

'-------------------------------------------------------------------------------
' BoundedTargetRange
' Returns UsedRange unioned with the active window's VisibleRange, so the
' highlight always covers what's on screen plus wherever the user's data
' actually lives, without formatting the full 17-billion-cell grid.
'
' NOTE: on a sheet with sparse data spread far apart, or when scrolled deep
' into empty space beyond both UsedRange and the current viewport, the
' crosshair may not be visible until that area is scrolled into view. This
' is a deliberate performance trade-off - see docs/architecture.md.
'-------------------------------------------------------------------------------
Private Function BoundedTargetRange(ByVal ws As Worksheet) As Range

    On Error GoTo Fallback

    Dim result As Range
    Set result = ws.UsedRange

    Dim wn As Window
    For Each wn In ws.Parent.Windows
        If wn.Visible Then
            On Error Resume Next
            If Not wn.RangeSelection Is Nothing Then
                If wn.ActiveSheet.name = ws.name Then
                    Set result = Union(result, wn.VisibleRange)
                End If
            End If
            On Error GoTo Fallback
        End If
    Next wn

    Set BoundedTargetRange = result
    Exit Function

Fallback:
    On Error Resume Next
    Set BoundedTargetRange = ws.UsedRange
    On Error GoTo 0

End Function

'-------------------------------------------------------------------------------
' AddRule
' Adds a single expression-based conditional format to targetRange.
' Supports both fill (Interior) and border styles.
'-------------------------------------------------------------------------------
Private Sub AddRule(ByVal targetRange As Range, ByVal formula As String, ByVal colour As Long, _
                    Optional ByVal style As HighlightStyle = hsFill)

    Dim fc As FormatCondition
    Set fc = targetRange.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)

    ' Empirically verified on Excel 2024: FormatCondition.Borders does NOT
    ' accept .Weight (raises error 1004). If Weight is assigned before
    ' Color (as older versions of this add-in did), the error aborts the
    ' block before Color is set, so the border renders thin and black no
    ' matter what colour was requested. Correct order: Colour first,
    ' LineStyle second, Weight last (best-effort, tolerated failure).
    If style = hsFill Then
        fc.Interior.Color = colour
    Else
        On Error Resume Next
        fc.Borders.Color = colour
        fc.Borders.LineStyle = xlContinuous
        fc.Borders.Weight = xlThick
        On Error GoTo 0
    End If

    On Error Resume Next
    fc.StopIfTrue = False
    fc.SetFirstPriority
    On Error GoTo 0

End Sub

'-------------------------------------------------------------------------------
' AddIntersectionRule
' Description : The cursor-cell accent in Crosshair mode. Kept separate from
'               AddRule for two reasons:
'               1. StopIfTrue=True - the accent must win at the exact
'                  intersection cell. Without it, Excel's rule precedence
'                  lets the later row/column rules override the accent, which
'                  is exactly why Intersect looked broken before.
'               2. In Border style the accent also fills the cell, so the
'                  cursor cell is clearly visible even though the surrounding
'                  highlight is border-only.
' Parameters  : targetRange   - range the rule is applied to
'               formula        - the expression formula
'               accentColour   - the (already darkened) accent fill colour
'               style          - the active highlight style (fill or border)
'-------------------------------------------------------------------------------
Private Sub AddIntersectionRule(ByVal targetRange As Range, ByVal formula As String, _
                                ByVal accentColour As Long, ByVal style As HighlightStyle)

    Dim fc As FormatCondition
    Set fc = targetRange.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)

    On Error Resume Next
    fc.Interior.Color = accentColour
    If style = hsBorder Then
        fc.Borders.Color = accentColour
        fc.Borders.LineStyle = xlContinuous
        fc.Borders.Weight = xlThick
    End If
    fc.StopIfTrue = True
    fc.SetFirstPriority
    On Error GoTo 0

End Sub

'-------------------------------------------------------------------------------
' IntersectionAccentRGB
' Description : The colour for the crosshair intersection accent cell. If the
'               user has explicitly configured a custom intersection colour
'               (anything other than the factory default), honour it -
'               otherwise derive a clearly darker version of the current
'               highlight colour so the accent is always visibly stronger
'               than the row/column band. A fixed accent colour equal to the
'               highlight colour is exactly why the old behaviour looked
'               broken (both were yellow).
' Returns     : Long - BGR colour value for the accent cell
'-------------------------------------------------------------------------------
Private Function IntersectionAccentRGB() As Long
    On Error Resume Next
    Dim base As Long
    base = Settings.EffectiveRGB
    ' RGB_YELLOW doubles as the "no custom colour configured" sentinel, so an
    ' explicitly configured YELLOW accent is deliberately treated as unset
    ' and falls through to the darkened auto-accent. (Known limitation: you
    ' can't pick plain yellow as a custom intersection colour - pick a
    ' different shade, or clear it to get the auto-darkened one.)
    If Settings.IntersectionRGB <> RGB_YELLOW And Settings.IntersectionRGB <> base Then
        IntersectionAccentRGB = Settings.IntersectionRGB
    Else
        IntersectionAccentRGB = Utilities.DarkenColour(base, 55)
    End If
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' RemoveOurConditionalFormatting
' Description : Scans every CF rule on the sheet (ws.Cells.FormatConditions
'               returns the full set regardless of which sub-range each rule
'               was originally applied to) and deletes only the ones whose
'               formula references our defined names. Anything the user
'               added themselves is left completely untouched.
'-------------------------------------------------------------------------------
Public Sub RemoveOurConditionalFormatting(ByVal ws As Worksheet)
' NOTE: Made Public so RibbonCallbacks.OnExclude_Action can strip formatting
' from excluded workbooks.

    On Error GoTo ErrHandler

    Dim allConditions As FormatConditions
    Set allConditions = ws.Cells.FormatConditions

    Dim i As Long
    For i = allConditions.count To 1 Step -1
        Dim fc As FormatCondition
        Set fc = allConditions(i)

        If fc.Type = xlExpression Then
            Dim f As String
            f = fc.Formula1
            If InStr(1, f, NAME_ROW_PREFIX, vbTextCompare) > 0 _
               Or InStr(1, f, NAME_COL_PREFIX, vbTextCompare) > 0 Then
                allConditions(i).Delete
            End If
        End If
    Next i

    Exit Sub

ErrHandler:
    Logging.LogError "HighlightEngine.RemoveOurConditionalFormatting", Err.Number, Err.Description, ws.name

End Sub

'-------------------------------------------------------------------------------
' UpdateStatusBar
' Description : Shows the current highlighter mode and colour in the status
'               bar so the user can see the state at a glance. Cleared when
'               highlighting is off or the workbook is excluded.
'-------------------------------------------------------------------------------
Private Sub UpdateStatusBar(ByVal wb As Workbook)

    On Error Resume Next

    If wb Is Nothing Then
        Application.StatusBar = False   ' reset to default
        Exit Sub
    End If

    If Not Settings.enabled Or Settings.Mode = hmNone Then
        Application.StatusBar = False
        Exit Sub
    End If

    If Utilities.WorkbookIsExcluded(wb) Then
        Application.StatusBar = STATUS_BAR_PREFIX & "Excluded"
        Exit Sub
    End If

    Dim modeText As String
    Select Case Settings.Mode
        Case hmRow:       modeText = "Row"
        Case hmColumn:    modeText = "Col"
        Case hmCrosshair: modeText = "Crosshair"
    End Select

    Dim colourText As String
    colourText = ColourName(Settings.colour)

    Dim extras As String
    If Settings.HighlightStyle = hsBorder Then extras = " [Border]"
    If Settings.IntersectionEnabled Then extras = extras & " [Intersect]"
    If Settings.AnimatedEnabled Then extras = extras & " [Pulse]"

    Application.StatusBar = STATUS_BAR_PREFIX & modeText & " - " & colourText & extras

End Sub

'-------------------------------------------------------------------------------
' TriggerAnimationPulse
' Description : Briefly flashes the highlight by cycling through a lighter
'               tint and back. Uses Application.OnTime for the timer loop.
'               Capped to 3 iterations so it doesn't fight performance.
'
'               NOTE (fixed in 2.3.0): older versions formatted the target
'               time to "hh:mm:ss" before scheduling. That dropped the date
'               AND the sub-second precision, so OnTime often received a time
'               that had already passed and the pulse never visibly fired.
'               The schedule time is now passed as a full-precision Date
'               (Now + 0.12 seconds as a fraction of a day).
'-------------------------------------------------------------------------------
Private Sub TriggerAnimationPulse(ByVal ws As Worksheet)

    On Error Resume Next

    ' Bump the generation so any in-flight chain from a previous selection
    ' change stops at its next tick instead of fighting this new pulse.
    mPulseGeneration = mPulseGeneration + 1

    Dim pulseData As String
    pulseData = ws.Parent.name & PULSE_SEP & ws.name
    Logging.LogInfo "HighlightEngine.TriggerAnimationPulse", "Pulse started for " & pulseData
    SchedulePulseStep pulseData, 0, mPulseGeneration

End Sub

'-------------------------------------------------------------------------------
' SchedulePulseStep
' Schedules one pulse tick 0.12 seconds from now. Stores no intermediate
' state - each tick re-derives its sheet from pulseData and its generation
' is checked in PulseStep.
'-------------------------------------------------------------------------------
Private Sub SchedulePulseStep(ByVal pulseData As String, ByVal iteration As Integer, ByVal gen As Long)
    On Error Resume Next
    Application.OnTime EarliestTime:=Now + 0.12 / 86400, _
        Procedure:="'HighlightEngine.PulseStep """ & pulseData & """, " & iteration & ", " & gen & "'"
End Sub

'-------------------------------------------------------------------------------
' PulseStep
' Called by Application.OnTime to animate the highlight. Iteration 0-2
' cycles the colour, then restores the original (full rebuild, so per-mode
' colours and the intersection accent are all restored exactly).
'-------------------------------------------------------------------------------
Public Sub PulseStep(ByVal pulseData As String, ByVal iteration As Integer, ByVal gen As Long)

    On Error GoTo ErrHandler

    ' A newer pulse has started - this chain is stale, stop it now.
    If gen <> mPulseGeneration Then Exit Sub

    Dim parts() As String
    parts = Split(pulseData, PULSE_SEP)
    If UBound(parts) < 1 Then Exit Sub

    Dim wb As Workbook
    Set wb = Application.Workbooks(parts(0))
    If wb Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Set ws = wb.Worksheets(parts(1))
    If ws Is Nothing Then Exit Sub

    If iteration > 2 Then
        ' Animation complete - rebuild from settings so every rule (row,
        ' column and the accent) returns to its exact configured colour.
        RebuildConditionalFormatting ws
        Exit Sub
    End If

    ' Pulse: alternate between a lighter version and the original. Each rule
    ' is handled individually so per-mode colours survive the middle ticks.
    Dim fc As FormatCondition
    Dim i As Long
    For i = 1 To ws.Cells.FormatConditions.count
        Set fc = ws.Cells.FormatConditions(i)
        If fc.Type = xlExpression Then
            Dim f As String
            f = fc.Formula1
            If InStr(1, f, NAME_ROW_PREFIX, vbTextCompare) > 0 _
               Or InStr(1, f, NAME_COL_PREFIX, vbTextCompare) > 0 Then
                If iteration Mod 2 = 0 Then
                    ' Lighter tint: blend the rule's current colour with white.
                    ' Interior.Color is BGR (0x00BBGGRR): low byte is red,
                    ' high byte is blue - read in that order.
                    Dim col As Long, r As Long, g As Long, b As Long
                    col = fc.Interior.Color
                    r = col Mod 256
                    g = (col \ 256) Mod 256
                    b = (col \ 65536) Mod 256
                    fc.Interior.Color = RGB((r + 255) \ 2, (g + 255) \ 2, (b + 255) \ 2)
                Else
                    ' Restore the rule's own colour by its axis, so per-mode
                    ' colours don't get flattened to a single effective value.
                    If InStr(1, f, NAME_ROW_PREFIX, vbTextCompare) > 0 And _
                       InStr(1, f, NAME_COL_PREFIX, vbTextCompare) > 0 Then
                        ' Intersection accent rule.
                        fc.Interior.Color = IntersectionAccentRGB()
                    ElseIf InStr(1, f, NAME_COL_PREFIX, vbTextCompare) > 0 Then
                        fc.Interior.Color = Settings.EffectiveColRGB
                    Else
                        fc.Interior.Color = Settings.EffectiveRowRGB
                    End If
                End If
            End If
        End If
    Next i

    ' Schedule next step.
    SchedulePulseStep pulseData, iteration + 1, gen

    Exit Sub

ErrHandler:
    Logging.LogError "HighlightEngine.PulseStep", Err.Number, Err.Description

End Sub

'-------------------------------------------------------------------------------
' Tracker helpers
'-------------------------------------------------------------------------------
Private Function TrackerKey(ByVal wb As Workbook, ByVal ws As Worksheet) As String
    TrackerKey = wb.name & "|" & ws.CodeName
End Function

Private Function CurrentSignature() As String
    ' Include per-mode colours in the signature so changing them triggers
    ' a CF rebuild. When per-mode is off, both return the same value.
    CurrentSignature = ModeToString(Settings.Mode) & "|" & _
                       Settings.EffectiveRowRGB & "|" & Settings.EffectiveColRGB & "|" & _
                       Settings.HighlightStyle
End Function

Private Function SheetMatchesCurrentSignature(ByVal wb As Workbook, ByVal ws As Worksheet) As Boolean
    EnsureTrackers
    Dim key As String
    key = TrackerKey(wb, ws)
    If Not mTrackers.Exists(key) Then
        SheetMatchesCurrentSignature = False
    Else
        SheetMatchesCurrentSignature = (mTrackers(key) = CurrentSignature())
    End If
End Function

Private Sub TrackSheet(ByVal wb As Workbook, ByVal ws As Worksheet)
    EnsureTrackers
    mTrackers(TrackerKey(wb, ws)) = CurrentSignature()
End Sub

Public Sub UntrackSheet(ByVal wb As Workbook, ByVal ws As Worksheet)
' NOTE: Made Public so RibbonCallbacks.OnExcludeSheet_Action can untrack
' sheets that are being excluded.
    EnsureTrackers
    Dim key As String
    key = TrackerKey(wb, ws)
    If mTrackers.Exists(key) Then mTrackers.Remove key
End Sub

Private Function IsTracked(ByVal wb As Workbook, ByVal ws As Worksheet) As Boolean
    EnsureTrackers
    IsTracked = mTrackers.Exists(TrackerKey(wb, ws))
End Function
