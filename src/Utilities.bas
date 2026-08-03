Attribute VB_Name = "Utilities"
'===============================================================================
' Module    : Utilities
' Purpose   : Small stateless helper functions shared by the rest of the
'             add-in. Nothing in here should touch Settings or the ribbon
'             directly - keep it a leaf module so it stays easy to test.
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' WorkbookIsEligible
' Description : Decides whether a workbook should ever receive highlighting.
'               Excludes the add-in's own workbook and anything that isn't a
'               normal worksheet-backed workbook (e.g. a workbook currently
'               mid-way through being created with no sheets, which can
'               briefly happen during automation).
' Parameters  : wb - the workbook to test
' Returns     : Boolean
' Example     : If WorkbookIsEligible(ActiveWorkbook) Then ...
'-------------------------------------------------------------------------------
Public Function WorkbookIsEligible(ByVal wb As Workbook) As Boolean

    On Error GoTo Fail

    If wb Is Nothing Then GoTo Fail
    If wb.IsAddin Then GoTo Fail                  ' never highlight the add-in itself
    If wb.ProtectStructure Then
        ' Structure-protected workbooks can still be highlighted; sheet-level
        ' protection is handled separately in HighlightEngine.
    End If

    WorkbookIsEligible = True
    Exit Function

Fail:
    WorkbookIsEligible = False

End Function

'-------------------------------------------------------------------------------
' SheetIsEligible
' Description : Filters out sheet types that don't support conditional
'               formatting the way we need (chart sheets, etc.) or that are
'               protected in a way that would make our CF calls raise 1004.
' Parameters  : sh - the sheet reported by a SheetXxx event (declared as Object
'                     because Application events pass chart sheets too)
' Returns     : Boolean
'-------------------------------------------------------------------------------
Public Function SheetIsEligible(ByVal sh As Object) As Boolean

    On Error GoTo Fail

    If sh Is Nothing Then GoTo Fail
    If Not TypeOf sh Is Worksheet Then GoTo Fail   ' skip chart sheets

    Dim ws As Worksheet
    Set ws = sh

    If ws.ProtectContents Then
        ' A protected sheet will usually reject FormatConditions.Add unless
        ' "Format cells" was left enabled when the sheet was protected. We
        ' still try - HighlightEngine traps and logs the 1004 if it happens -
        ' but flag it here so callers can short-circuit cheaply.
        SheetIsEligible = SheetAllowsFormattingWhenProtected(ws)
        Exit Function
    End If

    SheetIsEligible = True
    Exit Function

Fail:
    SheetIsEligible = False

End Function

'-------------------------------------------------------------------------------
' SheetAllowsFormattingWhenProtected
' Protection settings expose which operations remain allowed. There isn't a
' direct "AllowFormattingCells" read-back for conditional formatting itself,
' so we use AllowFormattingCells as the closest practical proxy.
'-------------------------------------------------------------------------------
Private Function SheetAllowsFormattingWhenProtected(ByVal ws As Worksheet) As Boolean
    On Error GoTo Fail
    SheetAllowsFormattingWhenProtected = ws.Protection.AllowFormattingCells
    Exit Function
Fail:
    SheetAllowsFormattingWhenProtected = False
End Function

'-------------------------------------------------------------------------------
' SafeDeleteName
' Description : Deletes a workbook-scoped defined name if it exists, silently
'               no-oping otherwise. Centralised so cleanup code never needs
'               its own On Error handler for the "name doesn't exist" case.
' Parameters  : wb - workbook that owns the name
'               nm - the name text (without workbook qualification)
'-------------------------------------------------------------------------------
Public Sub SafeDeleteName(ByVal wb As Workbook, ByVal nm As String)
    On Error Resume Next
    wb.Names(nm).Delete
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' WorkbookIsExcluded
' Description : Checks whether a workbook has the exclusion marker defined
'               name set to 1. The exclusion travels with the file itself
'               rather than living in the registry, so it survives being
'               sent to someone else.
'               Also honours the legacy pre-2.1.4 _XLCH_Excluded name so an
'               exclusion set by an older build keeps working after upgrade.
' Parameters  : wb - the workbook to check
' Returns     : Boolean - True if the workbook is excluded from highlighting
'-------------------------------------------------------------------------------
Public Function WorkbookIsExcluded(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    If NameExists(wb, NAME_EXCLUDED) Then
        WorkbookIsExcluded = (CLng(wb.Names(NAME_EXCLUDED).RefersToRange.value) = 1)
    ElseIf NameExists(wb, NAME_LEGACY_EXCLUDED) Then
        WorkbookIsExcluded = (CLng(wb.Names(NAME_LEGACY_EXCLUDED).RefersToRange.value) = 1)
    End If
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' SetWorkbookExclusion
' Description : Adds or removes the exclusion marker on a workbook. When
'               excluded is True, a hidden defined name XLCH_Excluded is
'               created (or updated) with value 1. When False, the name is
'               removed entirely. (Clean prefix, no leading underscore -
'               modern Excel 2024 rejects _XLCH_* names with error 1004.)
' Parameters  : wb       - the workbook to modify
'               excluded - True to exclude, False to allow highlighting
'-------------------------------------------------------------------------------
Public Sub SetWorkbookExclusion(ByVal wb As Workbook, ByVal excluded As Boolean)
    On Error GoTo ErrHandler
    If excluded Then
        If Not NameExists(wb, NAME_EXCLUDED) Then
            wb.Names.Add name:=NAME_EXCLUDED, RefersToR1C1:="=1", Visible:=False
        Else
            wb.Names(NAME_EXCLUDED).RefersToR1C1 = "=1"
        End If
    Else
        SafeDeleteName wb, NAME_EXCLUDED
    End If
    Exit Sub
ErrHandler:
    Logging.LogError "Utilities.SetWorkbookExclusion", Err.Number, Err.Description, wb.name
End Sub

'-------------------------------------------------------------------------------
' SheetIsExcluded
' Description : Checks whether a specific worksheet has the sheet-level
'               exclusion marker defined name set to 1. This is a
'               worksheet-scoped name, so it only affects that one sheet.
' Parameters  : ws - the worksheet to check
' Returns     : Boolean - True if the sheet is excluded from highlighting
'-------------------------------------------------------------------------------
Public Function SheetIsExcluded(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    If NameExists(ws.Parent, NAME_SHEET_EXCLUDED) Then
        ' Worksheet-scoped names are accessed via ws.Parent.Names but
        ' the name itself is scoped to the sheet. We check if the name
        ' exists and its value is 1.
        Dim nm As name
        Set nm = ws.Parent.Names(NAME_SHEET_EXCLUDED)
        If Not nm Is Nothing Then
            If InStr(1, nm.RefersTo, ws.CodeName, vbTextCompare) > 0 Then
                SheetIsExcluded = (CLng(ws.Parent.Names(NAME_SHEET_EXCLUDED).RefersToRange.value) = 1)
            End If
        End If
    ElseIf NameExists(ws.Parent, NAME_LEGACY_SHEET_EXCLUDED) Then
        ' Legacy pre-2.1.4 worksheet-scoped name - honour it so an exclusion
        ' set by an older build keeps working after upgrade.
        Set nm = ws.Parent.Names(NAME_LEGACY_SHEET_EXCLUDED)
        If Not nm Is Nothing Then
            If InStr(1, nm.RefersTo, ws.CodeName, vbTextCompare) > 0 Then
                SheetIsExcluded = (CLng(ws.Parent.Names(NAME_LEGACY_SHEET_EXCLUDED).RefersToRange.value) = 1)
            End If
        End If
    End If
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' SetSheetExclusion
' Description : Adds or removes the sheet-level exclusion marker. When
'               excluded is True, a worksheet-scoped defined name
'               XLCH_SheetExcluded is created with value 1. When False,
'               the name is removed.
' Parameters  : ws       - the worksheet to modify
'               excluded - True to exclude, False to allow highlighting
'-------------------------------------------------------------------------------
Public Sub SetSheetExclusion(ByVal ws As Worksheet, ByVal excluded As Boolean)
    On Error GoTo ErrHandler
    If excluded Then
        ' Worksheet-scoped name: the name includes the sheet name as scope.
        Dim scopedName As String
        scopedName = "'" & ws.name & "'!" & NAME_SHEET_EXCLUDED
        If Not NameExists(ws.Parent, NAME_SHEET_EXCLUDED) Then
            ws.Parent.Names.Add name:=scopedName, RefersToR1C1:="=1", Visible:=False
        Else
            ws.Parent.Names(NAME_SHEET_EXCLUDED).RefersToR1C1 = "=1"
        End If
    Else
        SafeDeleteName ws.Parent, NAME_SHEET_EXCLUDED
    End If
    Exit Sub
ErrHandler:
    Logging.LogError "Utilities.SetSheetExclusion", Err.Number, Err.Description, ws.name
End Sub

'-------------------------------------------------------------------------------
' NameExists
'-------------------------------------------------------------------------------
Public Function NameExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    On Error Resume Next
    NameExists = Not wb.Names(nm) Is Nothing
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' DeleteLegacyNames
' Description : Removes the pre-2.1.4 leading-underscore defined names
'               (_XLCH_*) from a workbook. Those names are rejected by modern
'               Excel parsers, and workbooks touched by older builds may still
'               carry them. Called from the cleanup paths (workbook close,
'               add-in off, exclusion toggle off) so upgrading never orphans
'               them and the "never leave a workbook modified" promise holds.
' Parameters  : wb - the workbook to sweep
'-------------------------------------------------------------------------------
Public Sub DeleteLegacyNames(ByVal wb As Workbook)
    On Error Resume Next
    SafeDeleteName wb, NAME_LEGACY_ROW_PREFIX
    SafeDeleteName wb, NAME_LEGACY_ROW_END_PREFIX
    SafeDeleteName wb, NAME_LEGACY_COL_PREFIX
    SafeDeleteName wb, NAME_LEGACY_COL_END_PREFIX
    SafeDeleteName wb, NAME_LEGACY_EXCLUDED
    SafeDeleteName wb, NAME_LEGACY_SHEET_EXCLUDED
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' ColourToRGB
' Description : Maps the HighlightColour enum to its RGB Long, falling back to
'               a persisted custom RGB when hcCustom is selected.
' Parameters  : colour    - the enum value
'               customRGB - value to return when colour = hcCustom
' Returns     : Long (RGB colour value)
'-------------------------------------------------------------------------------
Public Function ColourToRGB(ByVal colour As HighlightColour, ByVal CustomRGB As Long) As Long

    Select Case colour
        Case hcYellow: ColourToRGB = RGB_YELLOW
        Case hcGreen:  ColourToRGB = RGB_GREEN
        Case hcOrange: ColourToRGB = RGB_ORANGE
        Case hcCyan:   ColourToRGB = RGB_CYAN
        Case hcBlue:   ColourToRGB = RGB_BLUE
        Case hcPink:   ColourToRGB = RGB_PINK
        Case hcGrey:   ColourToRGB = RGB_GREY
        Case hcCustom: ColourToRGB = CustomRGB
        Case Else:     ColourToRGB = RGB_YELLOW
    End Select

End Function

'-------------------------------------------------------------------------------
' DarkenColour
' Description : Returns a BGR colour scaled toward black by the given
'               percentage (percent=55 means 55% of the original brightness).
'               Used for the crosshair intersection accent so the cursor cell
'               is always visibly stronger than the row/column band, whatever
'               highlight colour is in use.
' Parameters  : bgrColour - Excel BGR Long (0x00BBGGRR)
'               percent   - brightness percentage to keep (1-100)
' Returns     : Long - darkened BGR colour
'-------------------------------------------------------------------------------
Public Function DarkenColour(ByVal bgrColour As Long, ByVal percent As Long) As Long
    Dim r As Long, g As Long, b As Long
    r = bgrColour Mod 256
    g = (bgrColour \ 256) Mod 256
    b = (bgrColour \ 65536) Mod 256
    r = (r * percent) \ 100
    g = (g * percent) \ 100
    b = (b * percent) \ 100
    DarkenColour = RGB(r, g, b)
End Function

'-------------------------------------------------------------------------------
' ColourName
' Human-readable label used by the ribbon getLabel callbacks and the log.
'-------------------------------------------------------------------------------
Public Function ColourName(ByVal colour As HighlightColour) As String
    Select Case colour
        Case hcYellow: ColourName = "Yellow"
        Case hcGreen:  ColourName = "Green"
        Case hcOrange: ColourName = "Orange"
        Case hcCyan:   ColourName = "Cyan"
        Case hcBlue:   ColourName = "Blue"
        Case hcPink:   ColourName = "Pink"
        Case hcGrey:   ColourName = "Grey"
        Case hcCustom: ColourName = "Custom"
        Case Else:     ColourName = "Yellow"
    End Select
End Function

'-------------------------------------------------------------------------------
' ModeFromString / ColourFromString
' Round-trip helpers for the SaveSetting-backed persistence layer, which only
' stores strings.
'-------------------------------------------------------------------------------
Public Function ModeFromString(ByVal s As String) As HighlightMode
    Select Case UCase$(s)
        Case "ROW":       ModeFromString = hmRow
        Case "COLUMN":    ModeFromString = hmColumn
        Case "CROSSHAIR": ModeFromString = hmCrosshair
        Case "CELL":      ModeFromString = hmCell
        Case Else:        ModeFromString = hmNone
    End Select
End Function

Public Function ModeToString(ByVal m As HighlightMode) As String
    Select Case m
        Case hmRow:       ModeToString = "ROW"
        Case hmColumn:    ModeToString = "COLUMN"
        Case hmCrosshair: ModeToString = "CROSSHAIR"
        Case hmCell:      ModeToString = "CELL"
        Case Else:        ModeToString = "NONE"
    End Select
End Function

Public Function ColourFromString(ByVal s As String) As HighlightColour
    Select Case UCase$(s)
        Case "YELLOW": ColourFromString = hcYellow
        Case "GREEN":  ColourFromString = hcGreen
        Case "ORANGE": ColourFromString = hcOrange
        Case "CYAN":   ColourFromString = hcCyan
        Case "BLUE":   ColourFromString = hcBlue
        Case "PINK":   ColourFromString = hcPink
        Case "GREY":   ColourFromString = hcGrey
        Case "CUSTOM": ColourFromString = hcCustom
        Case Else:     ColourFromString = hcYellow
    End Select
End Function
