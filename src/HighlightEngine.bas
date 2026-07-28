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
'===============================================================================
Option Explicit

' Tracks which worksheets currently have our CF applied, and with what
' mode+colour signature, so we don't rebuild CF on every SelectionChange.
' Key   : Workbook.Name & "|" & Worksheet.Name
' Value : "<mode>|<rgb>" signature string
Private mTrackers As Object

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

    If Not Settings.Enabled Then Exit Sub
    If Settings.Mode = hmNone Then Exit Sub
    If Not Utilities.SheetIsEligible(sh) Then Exit Sub
    ' TODO: Support a per-workbook / per-worksheet exclusion list here once
    ' that setting exists - see docs/future-features.md.

    Dim ws As Worksheet
    Set ws = sh

    Dim wb As Workbook
    Set wb = ws.Parent

    If Not Utilities.WorkbookIsEligible(wb) Then Exit Sub

    EnsureNamesExist wb

    ' Cheap part: just move the crosshair. Runs on every single selection
    ' change, so it must stay fast even on huge sheets.
    UpdatePositionNames wb, target.Row, target.Column

    ' Expensive part: only runs when this sheet hasn't been configured yet
    ' for the current mode/colour (first visit, or settings just changed).
    If Not SheetMatchesCurrentSignature(wb, ws) Then
        RebuildConditionalFormatting ws
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
    EnsureNamesExist wb
    Exit Sub
ErrHandler:
    Logging.LogError "HighlightEngine.HandleWorkbookOpen", Err.Number, Err.Description, wb.Name
End Sub

Public Sub HandleWorkbookActivate(ByVal wb As Workbook)
    On Error GoTo ErrHandler
    If Not Utilities.WorkbookIsEligible(wb) Then Exit Sub
    If Not Settings.Enabled Then Exit Sub
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
    Utilities.SafeDeleteName wb, NAME_COL_PREFIX

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
                Utilities.SafeDeleteName wb, NAME_COL_PREFIX
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
Private Sub EnsureNamesExist(ByVal wb As Workbook)

    If Not Utilities.NameExists(wb, NAME_ROW_PREFIX) Then
        wb.Names.Add Name:=NAME_ROW_PREFIX, RefersToR1C1:="=1", Visible:=False
    End If

    If Not Utilities.NameExists(wb, NAME_COL_PREFIX) Then
        wb.Names.Add Name:=NAME_COL_PREFIX, RefersToR1C1:="=1", Visible:=False
    End If

End Sub

'-------------------------------------------------------------------------------
' UpdatePositionNames
' The hot path - runs on every selection change. Just repoints two constant
' names, which is orders of magnitude cheaper than touching FormatConditions.
'-------------------------------------------------------------------------------
Private Sub UpdatePositionNames(ByVal wb As Workbook, ByVal r As Long, ByVal c As Long)
    wb.Names(NAME_ROW_PREFIX).RefersToR1C1 = "=" & r
    wb.Names(NAME_COL_PREFIX).RefersToR1C1 = "=" & c
End Sub

'-------------------------------------------------------------------------------
' RebuildConditionalFormatting
' Description : Strips any of our existing rules from the sheet and, if the
'               add-in is enabled, adds fresh rules matching the current
'               mode and colour. Bounded to UsedRange ∪ VisibleRange to keep
'               this affordable on very large worksheets.
'-------------------------------------------------------------------------------
Private Sub RebuildConditionalFormatting(ByVal ws As Worksheet)

    On Error GoTo ErrHandler

    RemoveOurConditionalFormatting ws

    If Not Settings.Enabled Or Settings.Mode = hmNone Then
        UntrackSheet ws.Parent, ws
        Exit Sub
    End If

    Dim targetRange As Range
    Set targetRange = BoundedTargetRange(ws)
    If targetRange Is Nothing Then Exit Sub

    Dim colour As Long
    colour = Settings.EffectiveRGB

    Select Case Settings.Mode
        Case hmRow
            AddRule targetRange, "=ROW()=" & NAME_ROW_PREFIX, colour
        Case hmColumn
            AddRule targetRange, "=COLUMN()=" & NAME_COL_PREFIX, colour
        Case hmCrosshair
            AddRule targetRange, "=ROW()=" & NAME_ROW_PREFIX, colour
            AddRule targetRange, "=COLUMN()=" & NAME_COL_PREFIX, colour
    End Select

    TrackSheet ws.Parent, ws

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
'-------------------------------------------------------------------------------
Private Sub AddRule(ByVal targetRange As Range, ByVal formula As String, ByVal colour As Long)

    Dim fc As FormatCondition
    Set fc = targetRange.FormatConditions.Add(Type:=xlExpression, Formula1:=formula)

    With fc
        .Interior.Color = colour
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
Private Sub RemoveOurConditionalFormatting(ByVal ws As Worksheet)

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
' Tracker helpers
'-------------------------------------------------------------------------------
Private Function TrackerKey(ByVal wb As Workbook, ByVal ws As Worksheet) As String
    TrackerKey = wb.Name & "|" & ws.CodeName
End Function

Private Function CurrentSignature() As String
    CurrentSignature = ModeToString(Settings.Mode) & "|" & Settings.EffectiveRGB
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

Private Sub UntrackSheet(ByVal wb As Workbook, ByVal ws As Worksheet)
    EnsureTrackers
    Dim key As String
    key = TrackerKey(wb, ws)
    If mTrackers.Exists(key) Then mTrackers.Remove key
End Sub

Private Function IsTracked(ByVal wb As Workbook, ByVal ws As Worksheet) As Boolean
    EnsureTrackers
    IsTracked = mTrackers.Exists(TrackerKey(wb, ws))
End Function
