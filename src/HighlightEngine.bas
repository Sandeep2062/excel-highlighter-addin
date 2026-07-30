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
'             names (_XLCH_Row / _XLCH_Col) that hold the active row/column as
'             plain numeric constants. Conditional formatting rules on the
'             worksheet reference those names (e.g. "=ROW()=_XLCH_Row"), so
'             on every selection change we only need to update two Name
'             values - a very cheap operation - rather than add/remove CF
'             rules on every keystroke. CF rules themselves are only
'             (re)built when a sheet is seen for the first time, or when the
'             user changes mode/colour/enabled state from the ribbon.
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

' Timer ID for the animation pulse, so we can cancel it if settings change.
Private mAnimationTimerID As String

'-------------------------------------------------------------------------------
' EnsureTrackers
'-------------------------------------------------------------------------------
Private Sub EnsureTrackers()
    If mTrackers Is Nothing Then
        Set mTrackers = CreateObject("Scripting.Dictionary")
    End If
End Sub

'-------------------------------------------------------------------------------
' HandleSelectionChange
' Description : Main entry point, called by EventApp on every
'               Application.SheetSelectionChange.
' Parameters  : sh     - the sheet reported by the event (may be a chart sheet)
'               target - the new selection
'-------------------------------------------------------------------------------
Public Sub HandleSelectionChange(ByVal sh As Object, ByVal target As Range)

    On Error GoTo ErrHandler

    If Not Settings.Enabled Then
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
    SelectionHistory.Push wb.Name, ws.Name, target.Row, target.Column

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
    Logging.LogError "HighlightEngine.HandleSelectionChange", Err.Number, Err.Description, sh.Name

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
    If Not Settings.Enabled Then Exit Sub
    If Utilities.WorkbookIsExcluded(wb) Then Exit Sub
    EnsureNamesExist wb
    Exit Sub
ErrHandler:
    Logging.LogError "HighlightEngine.HandleWorkbookOpen", Err.Number, Err.Description, wb.Name
End Sub

Public Sub HandleWorkbookActivate(ByVal wb As Workbook)
    On Error GoTo ErrHandler
    If Not Utilities.WorkbookIsEligible(wb) Then Exit Sub
    If Not Settings.Enabled Then Exit Sub
    If Utilities.WorkbookIsExcluded(wb) Then Exit Sub
    EnsureNamesExist wb
    Exit Sub
ErrHandler:
    Logging.LogError "HighlightEngine.HandleWorkbookActivate", Err.Number, Err.Description, wb.Name
End Sub

Public Sub HandleWindowActivate(ByVal wb As Workbook)
    ' Same shape as WorkbookActivate - kept as a separate entry point so the
    ' two event sources stay independently traceable in the log.
    On Error GoTo ErrHandler
    If Not Utilities.WorkbookIsEligible(wb) Then Exit Sub
    If Not Settings.Enabled Then Exit Sub
    If Utilities.WorkbookIsExcluded(wb) Then Exit Sub
    EnsureNamesExist wb
    Exit Sub
ErrHandler:
    Logging.LogError "HighlightEngine.HandleWindowActivate", Err.Number, Err.Description, wb.Name
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

    Exit Sub

ErrHandler:
    ' Don't block the close on a logging or cleanup failure.
    Logging.LogError "HighlightEngine.HandleWorkbookBeforeClose", Err.Number, Err.Description, wb.Name

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

            If Not Settings.Enabled Then
                ' Turning the add-in off: strip everything back out.
                For Each ws In wb.Worksheets
                    RemoveOurConditionalFormatting ws
                    UntrackSheet wb, ws
                Next ws
                Utilities.SafeDeleteName wb, NAME_ROW_PREFIX
                Utilities.SafeDeleteName wb, NAME_ROW_END_PREFIX
                Utilities.SafeDeleteName wb, NAME_COL_PREFIX
                Utilities.SafeDeleteName wb, NAME_COL_END_PREFIX
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
        wb.Names.Add Name:=NAME_ROW_PREFIX, RefersToR1C1:="=1", Visible:=False
    End If

    If Not Utilities.NameExists(wb, NAME_ROW_END_PREFIX) Then
        wb.Names.Add Name:=NAME_ROW_END_PREFIX, RefersToR1C1:="=1", Visible:=False
    End If

    If Not Utilities.NameExists(wb, NAME_COL_PREFIX) Then
        wb.Names.Add Name:=NAME_COL_PREFIX, RefersToR1C1:="=1", Visible:=False
    End If

    If Not Utilities.NameExists(wb, NAME_COL_END_PREFIX) Then
        wb.Names.Add Name:=NAME_COL_END_PREFIX, RefersToR1C1:="=1", Visible:=False
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
        rEnd = ma.Row + ma.Rows.Count - 1
        cStart = ma.Column
        cEnd = ma.Column + ma.Columns.Count - 1
    Else
        rStart = target.Row
        rEnd = target.Row + target.Rows.Count - 1
        cStart = target.Column
        cEnd = target.Column + target.Columns.Count - 1
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

    If Not Settings.Enabled Or Settings.Mode = hmNone Then
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
            ' Add intersection accent if enabled.
            If Settings.IntersectionEnabled Then
                Dim interColour As Long
                interColour = Settings.IntersectionRGB
                AddRule targetRange, "=AND(" & rowExpr & "," & colExpr & ")", interColour, style
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
    Logging.LogError "HighlightEngine.RebuildConditionalFormatting", Err.Number, Err.Description, ws.Name

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
                If wn.ActiveSheet.Name = ws.Name Then
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

    With fc
        If style = hsFill Then
            .Interior.Color = colour
        Else
            ' Border style: use a thick border in the given colour.
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThick
            .Borders.Color = colour
        End If
        .StopIfTrue = False
        On Error Resume Next
        .SetFirstPriority
        On Error GoTo 0
    End With

End Sub

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
    For i = allConditions.Count To 1 Step -1
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
    Logging.LogError "HighlightEngine.RemoveOurConditionalFormatting", Err.Number, Err.Description, ws.Name

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

    If Not Settings.Enabled Or Settings.Mode = hmNone Then
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
    colourText = ColourName(Settings.Colour)

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
'-------------------------------------------------------------------------------
Private Sub TriggerAnimationPulse(ByVal ws As Worksheet)

    On Error Resume Next

    ' Cancel any existing animation timer.
    If Len(mAnimationTimerID) > 0 Then
        On Error Resume Next
        Application.OnTime EarliestTime:=CDate(mAnimationTimerID), Procedure:="", Schedule:=False
        On Error GoTo 0
    End If

    ' Schedule the first pulse step.
    mAnimationTimerID = Format$(Now + TimeValue("00:00:00.1"), "hh:mm:ss")
    ' Store the sheet info for the pulse callback.
    Dim pulseData As String
    pulseData = ws.Parent.Name & "|" & ws.Name
    Application.OnTime EarliestTime:=CDate(mAnimationTimerID), _
        Procedure:="'HighlightEngine.PulseStep """ & pulseData & """, 0'"

End Sub

'-------------------------------------------------------------------------------
' PulseStep
' Called by Application.OnTime to animate the highlight. Iteration 0-2
' cycles the colour, then restores the original.
'-------------------------------------------------------------------------------
Public Sub PulseStep(ByVal pulseData As String, ByVal iteration As Integer)

    On Error GoTo ErrHandler

    If iteration > 2 Then
        ' Animation complete - restore original colour.
        Dim parts() As String
        parts = Split(pulseData, "|")
        If UBound(parts) >= 1 Then
            Dim wb As Workbook
            Set wb = Application.Workbooks(parts(0))
            If Not wb Is Nothing Then
                Dim ws As Worksheet
                Set ws = wb.Worksheets(parts(1))
                If Not ws Is Nothing Then
                    RebuildConditionalFormatting ws
                End If
            End If
        End If
        mAnimationTimerID = ""
        Exit Sub
    End If

    ' Pulse: alternate between a lighter version and the original.
    parts = Split(pulseData, "|")
    If UBound(parts) >= 1 Then
        Set wb = Application.Workbooks(parts(0))
        If Not wb Is Nothing Then
            Set ws = wb.Worksheets(parts(1))
            If Not ws Is Nothing Then
                Dim fc As FormatCondition
                Dim i As Long
                For i = 1 To ws.Cells.FormatConditions.Count
                    Set fc = ws.Cells.FormatConditions(i)
                    If fc.Type = xlExpression Then
                        Dim f As String
                        f = fc.Formula1
                        If InStr(1, f, NAME_ROW_PREFIX, vbTextCompare) > 0 _
                           Or InStr(1, f, NAME_COL_PREFIX, vbTextCompare) > 0 Then
                            If iteration Mod 2 = 0 Then
                                ' Lighter tint: blend with white.
                                Dim r As Long, g As Long, b As Long
                                r = (fc.Interior.Color \ 65536) Mod 256
                                g = (fc.Interior.Color \ 256) Mod 256
                                b = fc.Interior.Color Mod 256
                                fc.Interior.Color = RGB( _
                                    (r + 255) \ 2, (g + 255) \ 2, (b + 255) \ 2)
                            Else
                                ' Restore original.
                                fc.Interior.Color = Settings.EffectiveRGB
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    End If

    ' Schedule next step.
    mAnimationTimerID = Format$(Now + TimeValue("00:00:00.15"), "hh:mm:ss")
    Application.OnTime EarliestTime:=CDate(mAnimationTimerID), _
        Procedure:="'HighlightEngine.PulseStep """ & pulseData & """, " & (iteration + 1) & "'"

    Exit Sub

ErrHandler:
    Logging.LogError "HighlightEngine.PulseStep", Err.Number, Err.Description
    mAnimationTimerID = ""

End Sub

'-------------------------------------------------------------------------------
' Tracker helpers
'-------------------------------------------------------------------------------
Private Function TrackerKey(ByVal wb As Workbook, ByVal ws As Worksheet) As String
    TrackerKey = wb.Name & "|" & ws.CodeName
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