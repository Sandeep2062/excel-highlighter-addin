Attribute VB_Name = "RibbonCallbacks"
'===============================================================================
' Module    : RibbonCallbacks
' Purpose   : Every procedure referenced from customUI14.xml lives here. Kept
'             thin - callbacks read/write Settings and HighlightEngine, they
'             never contain highlighting logic themselves.
'
'             Image loading: the colour gallery uses real PNG swatches
'             shipped in the "images" folder next to the .xlam (see
'             docs/installation.md). getImage callbacks load them with
'             LoadPicture, which is the simplest reliable way to supply
'             custom ribbon images without building an OPC image part.
'===============================================================================
Option Explicit

Private mRibbon As IRibbonUI

'-------------------------------------------------------------------------------
' onLoad
' RibbonX calls this once when the ribbon is built, passing the IRibbonUI
' we need to call Invalidate on later.
'-------------------------------------------------------------------------------
Public Sub onLoad(ByVal ribbon As IRibbonUI)
    On Error GoTo ErrHandler
    Set mRibbon = ribbon
    Logging.LogInfo "RibbonCallbacks.onLoad", "Ribbon UI attached"
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.onLoad", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' InvalidateRibbon
' Central place to refresh the ribbon after any state change, so every
' getPressed/getEnabled/getLabel callback gets re-queried.
'-------------------------------------------------------------------------------
Private Sub InvalidateRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
End Sub

'===============================================================================
' Enable / Disable toggle
'===============================================================================

Public Sub OnToggle_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    Settings.Enabled = pressed
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnToggle_Action", Err.Number, Err.Description
End Sub

Public Sub GetToggle_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    pressed = Settings.Enabled
End Sub

Public Sub GetToggle_Label(ByVal control As IRibbonControl, ByRef label)
    label = IIf(Settings.Enabled, "Highlight: On", "Highlight: Off")
End Sub

'===============================================================================
' Mode buttons (Row / Column / Crosshair) - behave as a manual radio group
'===============================================================================

Public Sub OnMode_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)

    On Error GoTo ErrHandler

    Dim requestedMode As HighlightMode
    requestedMode = ModeForControlId(control.ID)

    If pressed Then
        Settings.Mode = requestedMode
    ElseIf Settings.Mode = requestedMode Then
        ' User clicked the already-active mode button to switch highlighting off.
        Settings.Mode = hmNone
    End If

    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon

    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnMode_Action", Err.Number, Err.Description
End Sub

Public Sub GetMode_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    pressed = (Settings.Mode = ModeForControlId(control.ID))
End Sub

Public Sub GetMode_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.Enabled
End Sub

Private Function ModeForControlId(ByVal controlId As String) As HighlightMode
    Select Case controlId
        Case CTRL_MODE_ROW:       ModeForControlId = hmRow
        Case CTRL_MODE_COLUMN:    ModeForControlId = hmColumn
        Case CTRL_MODE_CROSSHAIR: ModeForControlId = hmCrosshair
        Case Else:                ModeForControlId = hmNone
    End Select
End Function

'===============================================================================
' Colour gallery
'===============================================================================

Public Sub OnGallery_Action(ByVal control As IRibbonControl, ByVal id As String, ByVal index As Integer)
    On Error GoTo ErrHandler
    Settings.Colour = ColourFromString(Mid$(id, Len("item") + 1))
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnGallery_Action", Err.Number, Err.Description
End Sub

Public Sub GetGallery_SelectedItemID(ByVal control As IRibbonControl, ByRef id)
    id = "item" & ColourName(Settings.Colour)
End Sub

Public Sub GetGallery_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.Enabled
End Sub

'-------------------------------------------------------------------------------
' GetSwatchImage
' Shared getImage handler for every gallery <item> and the custom colour
' button. The control id follows the "item<ColourName>" convention defined
' in customUI14.xml, e.g. "itemYellow" -> images\yellow.png.
'-------------------------------------------------------------------------------
Public Sub GetSwatchImage(ByVal control As IRibbonControl, ByRef image)

    On Error GoTo ErrHandler

    Dim colourPart As String
    If control.ID = CTRL_CUSTOM_COLOUR Then
        colourPart = "custom"
    Else
        colourPart = LCase$(Mid$(control.ID, Len("item") + 1))
    End If

    Dim path As String
    path = ImagesFolder() & colourPart & ".png"

    If Len(Dir$(path)) > 0 Then
        Set image = LoadPicture(path)
    End If

    Exit Sub

ErrHandler:
    ' Missing/unreadable image file - the gallery item just falls back to a
    ' blank icon rather than breaking ribbon load.
    Logging.LogError "RibbonCallbacks.GetSwatchImage", Err.Number, Err.Description, control.ID

End Sub

Private Function ImagesFolder() As String
    ImagesFolder = ThisWorkbook.Path & Application.PathSeparator & "images" & Application.PathSeparator
End Function

'===============================================================================
' Custom colour
'===============================================================================

Public Sub OnCustomColour_Action(ByVal control As IRibbonControl)

    On Error GoTo ErrHandler

    Dim chosen As Long
    If PromptForCustomRGB(chosen) Then
        Settings.CustomRGB = chosen
        Settings.Colour = hcCustom
        HighlightEngine.ReapplyAllOpenWorkbooks
        InvalidateRibbon
    End If

    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnCustomColour_Action", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' PromptForCustomRGB
' Description : Uses Excel's own colour picker (the small standard "Colors"
'               dialog) via Application.Dialogs so the user gets a real
'               palette/custom tab instead of typing numbers.
' Returns     : Boolean - True if the user picked a colour (OK, not Cancel)
' Parameters  : result - ByRef Long, the chosen RGB value when True is returned
'-------------------------------------------------------------------------------
Private Function PromptForCustomRGB(ByRef result As Long) As Boolean

    On Error GoTo Fallback

    ' xlDialogEditColor expects a palette index (1-56) to seed/replace and
    ' returns True/False via .Show; the resulting colour has to be read back
    ' from the workbook's colour palette entry we asked it to edit. We use a
    ' scratch palette slot (56, the last one) so we don't disturb any colour
    ' the user's workbook is actively using elsewhere.
    Const SCRATCH_PALETTE_INDEX As Long = 56

    Dim originalRGB As Long
    originalRGB = ActiveWorkbook.Colors(SCRATCH_PALETTE_INDEX)

    If Application.Dialogs(xlDialogEditColor).Show(SCRATCH_PALETTE_INDEX, Settings.CustomRGB) Then
        result = ActiveWorkbook.Colors(SCRATCH_PALETTE_INDEX)
        ActiveWorkbook.Colors(SCRATCH_PALETTE_INDEX) = originalRGB   ' restore the slot
        PromptForCustomRGB = True
    Else
        PromptForCustomRGB = False
    End If

    Exit Function

Fallback:
    ' No active workbook, or the dialog is unavailable in this context
    ' (e.g. called with no workbooks open) - degrade to a simple prompt
    ' rather than failing silently.
    Logging.LogError "RibbonCallbacks.PromptForCustomRGB", Err.Number, Err.Description
    PromptForCustomRGB = False

End Function

'===============================================================================
' Reset settings
'===============================================================================

Public Sub OnReset_Action(ByVal control As IRibbonControl)
    On Error GoTo ErrHandler
    Settings.ResetToDefaults
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnReset_Action", Err.Number, Err.Description
End Sub

'===============================================================================
' About
'===============================================================================

Public Sub OnAbout_Action(ByVal control As IRibbonControl)
    On Error Resume Next
    MsgBox "Excel Crosshair Highlighter" & vbCrLf & _
           "Version " & APP_VERSION & vbCrLf & vbCrLf & _
           "Non-destructive row/column/crosshair highlighting for the active cell." & vbCrLf & _
           "https://github.com/Sandeep2062/excel-highlighter-addin", _
           vbInformation, "About"
End Sub
