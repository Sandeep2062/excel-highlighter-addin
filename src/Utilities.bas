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
' NameExists
'-------------------------------------------------------------------------------
Public Function NameExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    On Error Resume Next
    NameExists = Not wb.Names(nm) Is Nothing
    On Error GoTo 0
End Function

'-------------------------------------------------------------------------------
' ColourToRGB
' Description : Maps the HighlightColour enum to its RGB Long, falling back to
'               a persisted custom RGB when hcCustom is selected.
' Parameters  : colour    - the enum value
'               customRGB - value to return when colour = hcCustom
' Returns     : Long (RGB colour value)
'-------------------------------------------------------------------------------
Public Function ColourToRGB(ByVal colour As HighlightColour, ByVal customRGB As Long) As Long

    Select Case colour
        Case hcYellow: ColourToRGB = RGB_YELLOW
        Case hcGreen:  ColourToRGB = RGB_GREEN
        Case hcOrange: ColourToRGB = RGB_ORANGE
        Case hcCyan:   ColourToRGB = RGB_CYAN
        Case hcBlue:   ColourToRGB = RGB_BLUE
        Case hcPink:   ColourToRGB = RGB_PINK
        Case hcGrey:   ColourToRGB = RGB_GREY
        Case hcCustom: ColourToRGB = customRGB
        Case Else:     ColourToRGB = RGB_YELLOW
    End Select

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
        Case Else:        ModeFromString = hmNone
    End Select
End Function

Public Function ModeToString(ByVal m As HighlightMode) As String
    Select Case m
        Case hmRow:       ModeToString = "ROW"
        Case hmColumn:    ModeToString = "COLUMN"
        Case hmCrosshair: ModeToString = "CROSSHAIR"
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
