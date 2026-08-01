Attribute VB_Name = "RibbonCallbacks"
'===============================================================================
' Module    : RibbonCallbacks
' Purpose   : Every procedure referenced from customUI14.xml lives here. Kept
'             thin - callbacks read/write Settings and HighlightEngine, they
'             never contain highlighting logic themselves.
'
'             Image loading: colour swatches are generated in memory as
'             solid-colour bitmaps via ColourPicker.CreateDynamicColourSwatch
'             (GDI + OleCreatePictureIndirect). No external PNG files are
'             needed - every gallery item and the custom colour button get
'             their image from a single shared helper.
'===============================================================================
Option Explicit

Private mRibbon As IRibbonUI

' The number of preset colours in the gallery.
Private Const PRESET_COLOUR_COUNT As Long = 7

'-------------------------------------------------------------------------------
' onLoad
' RibbonX calls this once when the ribbon is built, passing the IRibbonUI
' we need to call Invalidate on later.
'-------------------------------------------------------------------------------
Public Sub onLoad(ByVal ribbon As IRibbonUI)
    On Error GoTo ErrHandler
    Set mRibbon = ribbon
    Profiles.Init
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

'-------------------------------------------------------------------------------
' InvalidateRibbonExternally
' Public entry point so AddinHost (which doesn't have access to mRibbon) can
' trigger a ribbon refresh after the hotkey toggle fires.
'-------------------------------------------------------------------------------
Public Sub InvalidateRibbonExternally()
    InvalidateRibbon
End Sub

'===============================================================================
' Enable / Disable toggle
'===============================================================================

Public Sub OnToggle_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    Settings.enabled = pressed
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnToggle_Action", Err.Number, Err.Description
End Sub

Public Sub GetToggle_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    pressed = Settings.enabled
End Sub

Public Sub GetToggle_Label(ByVal control As IRibbonControl, ByRef label)
    label = IIf(Settings.enabled, "Highlight: On", "Highlight: Off")
End Sub

'===============================================================================
' Mode buttons (Row / Column / Crosshair) - behave as a manual radio group
'===============================================================================

Public Sub OnMode_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)

    On Error GoTo ErrHandler

    Dim requestedMode As HighlightMode
    requestedMode = ModeForControlId(control.id)

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
    pressed = (Settings.Mode = ModeForControlId(control.id))
End Sub

Public Sub GetMode_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.enabled
End Sub

Private Function ModeForControlId(ByVal controlId As String) As HighlightMode
    Select Case controlId
        Case CTRL_MODE_ROW:       ModeForControlId = hmRow
        Case CTRL_MODE_COLUMN:    ModeForControlId = hmColumn
        Case CTRL_MODE_CROSSHAIR: ModeForControlId = hmCrosshair
        Case CTRL_MODE_CELL:      ModeForControlId = hmCell
        Case Else:                ModeForControlId = hmNone
    End Select
End Function

'===============================================================================
' Colour gallery (dynamic)
' The gallery has two sections:
'   1. 7 preset colours (items 0-6, IDs "itemYellow" through "itemGrey")
'   2. Recent custom colours (items 7+, IDs "recent0" through "recent3")
' Total item count = 7 + recentCount
'===============================================================================

Public Sub OnGallery_Action(ByVal control As IRibbonControl, ByVal id As String, ByVal index As Integer)

    On Error GoTo ErrHandler

    If Left$(id, 4) = "item" Then
        ' Preset colour selected.
        Settings.colour = ColourFromString(Mid$(id, 5))
    ElseIf Left$(id, 6) = "recent" Then
        ' Recent colour selected. Apply it as custom.
        Dim recentIndex As Long
        recentIndex = CLng(Mid$(id, 7))
        If recentIndex >= 0 And recentIndex < Settings.RecentColourCount Then
            Settings.CustomRGB = Settings.RecentColour(recentIndex)
            Settings.colour = hcCustom
        End If
    End If

    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon

    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnGallery_Action", Err.Number, Err.Description
End Sub

Public Sub GetGallery_SelectedItemID(ByVal control As IRibbonControl, ByRef id)
    If Settings.colour = hcCustom Then
        ' Custom colour - select Nothing (no preset), or highlight the
        ' first recent colour that matches.
        id = ""
    Else
        id = "item" & ColourName(Settings.colour)
    End If
End Sub

Public Sub GetGallery_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.enabled
End Sub

'-------------------------------------------------------------------------------
' getItemCount - returns total number of dynamic gallery items
'-------------------------------------------------------------------------------
Public Sub GetGallery_ItemCount(ByVal control As IRibbonControl, ByRef count)
    count = PRESET_COLOUR_COUNT + Settings.RecentColourCount
End Sub

'-------------------------------------------------------------------------------
' getItemID
'-------------------------------------------------------------------------------
Public Sub GetGallery_ItemID(ByVal control As IRibbonControl, ByVal index As Integer, ByRef id)
    If index < PRESET_COLOUR_COUNT Then
        id = "item" & PresetColourName(index)
    Else
        id = "recent" & (index - PRESET_COLOUR_COUNT)
    End If
End Sub

'-------------------------------------------------------------------------------
' getItemLabel
'-------------------------------------------------------------------------------
Public Sub GetGallery_ItemLabel(ByVal control As IRibbonControl, ByVal index As Integer, ByRef label)
    If index < PRESET_COLOUR_COUNT Then
        label = PresetColourName(index)
    Else
        label = "Recent " & (index - PRESET_COLOUR_COUNT + 1)
    End If
End Sub

'-------------------------------------------------------------------------------
' getItemImage
' Generates a solid-colour bitmap swatch in memory for every gallery item,
' whether preset or recent. No disk I/O needed.
'-------------------------------------------------------------------------------
Public Sub GetGallery_ItemImage(ByVal control As IRibbonControl, ByVal index As Integer, ByRef image)

    On Error GoTo ErrHandler

    If index < PRESET_COLOUR_COUNT Then
        Set image = GenerateColourSwatch(PresetColourRGB(index))
    Else
        Dim recentIndex As Long
        recentIndex = index - PRESET_COLOUR_COUNT
        If recentIndex >= 0 And recentIndex < Settings.RecentColourCount Then
            Set image = GenerateColourSwatch(Settings.RecentColour(recentIndex))
        End If
    End If

    Exit Sub

ErrHandler:
    Logging.LogError "RibbonCallbacks.GetGallery_ItemImage", Err.Number, Err.Description, "index=" & index
End Sub

'-------------------------------------------------------------------------------
' getItemScreentip
'-------------------------------------------------------------------------------
Public Sub GetGallery_ItemScreentip(ByVal control As IRibbonControl, ByVal index As Integer, ByRef screentip)
    If index < PRESET_COLOUR_COUNT Then
        screentip = PresetColourName(index)
    Else
        screentip = "Recent custom colour"
    End If
End Sub

'-------------------------------------------------------------------------------
' PresetColourName
' Maps a 0-based index to the preset colour name string.
'-------------------------------------------------------------------------------
Private Function PresetColourName(ByVal index As Integer) As String
    Select Case index
        Case 0: PresetColourName = "Yellow"
        Case 1: PresetColourName = "Green"
        Case 2: PresetColourName = "Orange"
        Case 3: PresetColourName = "Cyan"
        Case 4: PresetColourName = "Blue"
        Case 5: PresetColourName = "Pink"
        Case 6: PresetColourName = "Grey"
        Case Else: PresetColourName = "Yellow"
    End Select
End Function

Private Function PresetColourRGB(ByVal index As Integer) As Long
    Select Case index
        Case 0: PresetColourRGB = RGB_YELLOW
        Case 1: PresetColourRGB = RGB_GREEN
        Case 2: PresetColourRGB = RGB_ORANGE
        Case 3: PresetColourRGB = RGB_CYAN
        Case 4: PresetColourRGB = RGB_BLUE
        Case 5: PresetColourRGB = RGB_PINK
        Case 6: PresetColourRGB = RGB_GREY
        Case Else: PresetColourRGB = RGB_YELLOW
    End Select
End Function

'-------------------------------------------------------------------------------
' GetSwatchImage
' Shared getImage handler for the custom colour button.
'-------------------------------------------------------------------------------
Public Sub GetSwatchImage(ByVal control As IRibbonControl, ByRef image)

    On Error GoTo ErrHandler

    Set image = GenerateColourSwatch(Settings.CustomRGB)

    Exit Sub

ErrHandler:
    Logging.LogError "RibbonCallbacks.GetSwatchImage", Err.Number, Err.Description, control.id

End Sub

'-------------------------------------------------------------------------------
' GenerateColourSwatch
' Creates a small solid-colour bitmap at runtime for recent colour swatches.
' Uses GDI and OleCreatePictureIndirect via ColourPicker module.
' Returns an IPictureDisp that the ribbon displays directly.
'-------------------------------------------------------------------------------
Private Function GenerateColourSwatch(ByVal rgbColour As Long) As Object
    On Error Resume Next
    Set GenerateColourSwatch = ColourPicker.CreateDynamicColourSwatch(rgbColour)
End Function

'===============================================================================
' Custom colour
'===============================================================================

Public Sub OnCustomColour_Action(ByVal control As IRibbonControl)

    On Error GoTo ErrHandler

    Dim chosen As Long
    ' Use the native Windows colour picker (ColourPicker.bas) instead of
    ' xlDialogEditColor, which borrowed palette slot 56 from the workbook.
    If ColourPicker.PromptForCustomRGB(chosen, Settings.CustomRGB) Then
        Settings.CustomRGB = chosen
        Settings.colour = hcCustom
        HighlightEngine.ReapplyAllOpenWorkbooks
        InvalidateRibbon
    End If

    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnCustomColour_Action", Err.Number, Err.Description
End Sub

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
' Workbook exclusion
'===============================================================================

Public Sub OnExcludeSheet_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo ErrHandler
    If ws Is Nothing Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    If Not Utilities.SheetIsEligible(ws) Then Exit Sub
    
    Utilities.SetSheetExclusion ws, pressed
    If pressed Then
        ' Strip highlighting from this sheet immediately.
        HighlightEngine.RemoveOurConditionalFormatting ws
        HighlightEngine.UntrackSheet ws.Parent, ws
        Logging.LogInfo "RibbonCallbacks.OnExcludeSheet_Action", "Excluded sheet " & ws.name
    Else
        ' Re-enable highlighting on this sheet.
        If Settings.enabled And Settings.Mode <> hmNone Then
            HighlightEngine.RebuildConditionalFormatting ws
        End If
        Logging.LogInfo "RibbonCallbacks.OnExcludeSheet_Action", "Un-excluded sheet " & ws.name
    End If
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnExcludeSheet_Action", Err.Number, Err.Description
End Sub

Public Sub GetExcludeSheet_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then
        pressed = False
        Exit Sub
    End If
    pressed = Utilities.SheetIsExcluded(ws)
End Sub

Public Sub GetExcludeSheet_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then
        enabled = False
        Exit Sub
    End If
    enabled = Utilities.SheetIsEligible(ws) And Not Utilities.WorkbookIsExcluded(ws.Parent)
End Sub

Public Sub GetExcludeSheet_Label(ByVal control As IRibbonControl, ByRef label)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ActiveSheet
    If ws Is Nothing Then
        label = "Exclude Sheet"
        Exit Sub
    End If
    If Utilities.SheetIsExcluded(ws) Then
        label = "Include Sheet"
    Else
        label = "Exclude Sheet"
    End If
End Sub

Public Sub OnExclude_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then Exit Sub
    If Not Utilities.WorkbookIsEligible(wb) Then Exit Sub
    Utilities.SetWorkbookExclusion wb, pressed
    If pressed Then
        ' Strip highlighting from this workbook immediately.
        Dim ws As Worksheet
        For Each ws In wb.Worksheets
            HighlightEngine.RemoveOurConditionalFormatting ws
        Next ws
        Utilities.SafeDeleteName wb, NAME_ROW_PREFIX
        Utilities.SafeDeleteName wb, NAME_COL_PREFIX
        Logging.LogInfo "RibbonCallbacks.OnExclude_Action", "Excluded " & wb.name
    Else
        ' Re-enable highlighting on this workbook.
        If Settings.enabled And Settings.Mode <> hmNone Then
            HighlightEngine.EnsureNamesExist wb
            On Error Resume Next
            If TypeOf wb.ActiveSheet Is Worksheet Then
                HighlightEngine.RebuildConditionalFormatting wb.ActiveSheet
            End If
            On Error GoTo ErrHandler
        End If
        Logging.LogInfo "RibbonCallbacks.OnExclude_Action", "Un-excluded " & wb.name
    End If
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnExclude_Action", Err.Number, Err.Description
End Sub

Public Sub GetExclude_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        pressed = False
        Exit Sub
    End If
    pressed = Utilities.WorkbookIsExcluded(wb)
End Sub

Public Sub GetExclude_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    On Error Resume Next
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        enabled = False
        Exit Sub
    End If
    enabled = Utilities.WorkbookIsEligible(wb)
End Sub

Public Sub GetExclude_Label(ByVal control As IRibbonControl, ByRef label)
    On Error Resume Next
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        label = "Exclude Workbook"
        Exit Sub
    End If
    If Utilities.WorkbookIsExcluded(wb) Then
        label = "Include Workbook"
    Else
        label = "Exclude Workbook"
    End If
End Sub

'===============================================================================
' Highlight style (Fill vs Border)
'===============================================================================

Public Sub OnStyle_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    Dim requestedStyle As HighlightStyle
    requestedStyle = StyleForControlId(control.id)
    Settings.HighlightStyle = requestedStyle
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnStyle_Action", Err.Number, Err.Description
End Sub

Public Sub GetStyle_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    pressed = (Settings.HighlightStyle = StyleForControlId(control.id))
End Sub

Public Sub GetStyle_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.enabled
End Sub

Private Function StyleForControlId(ByVal controlId As String) As HighlightStyle
    Select Case controlId
        Case CTRL_STYLE_FILL:   StyleForControlId = hsFill
        Case CTRL_STYLE_BORDER: StyleForControlId = hsBorder
        Case Else:              StyleForControlId = hsFill
    End Select
End Function

'===============================================================================
' Intersection
'===============================================================================

Public Sub OnIntersection_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    Settings.IntersectionEnabled = pressed
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnIntersection_Action", Err.Number, Err.Description
End Sub

Public Sub GetIntersection_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    pressed = Settings.IntersectionEnabled
End Sub

Public Sub GetIntersection_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.enabled
End Sub

'===============================================================================
' Protected
'===============================================================================

Public Sub OnAllowProtected_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    Settings.AllowProtected = pressed
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnAllowProtected_Action", Err.Number, Err.Description
End Sub

Public Sub GetAllowProtected_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    pressed = Settings.AllowProtected
End Sub

Public Sub GetAllowProtected_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.enabled
End Sub

'===============================================================================
' Animated
'===============================================================================

Public Sub OnAnimated_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    Settings.AnimatedEnabled = pressed
    ' No need to reapply - the pulse triggers on the next selection change.
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnAnimated_Action", Err.Number, Err.Description
End Sub

Public Sub GetAnimated_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    pressed = Settings.AnimatedEnabled
End Sub

Public Sub GetAnimated_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.enabled
End Sub

'===============================================================================
' Dark Mode
'===============================================================================

Public Sub OnDarkMode_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    Settings.DarkMode = pressed
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnDarkMode_Action", Err.Number, Err.Description
End Sub

Public Sub GetDarkMode_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    pressed = Settings.DarkMode
End Sub

Public Sub GetDarkMode_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.enabled
End Sub

'===============================================================================
' Selection history navigation
'===============================================================================

Public Sub OnHistoryBack_Action(ByVal control As IRibbonControl)
    On Error GoTo ErrHandler
    If SelectionHistory.CanGoBack Then
        SelectionHistory.HotkeyGoBack
    End If
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnHistoryBack_Action", Err.Number, Err.Description
End Sub

Public Sub GetHistoryBack_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = SelectionHistory.CanGoBack
End Sub

Public Sub OnHistoryForward_Action(ByVal control As IRibbonControl)
    On Error GoTo ErrHandler
    If SelectionHistory.CanGoForward Then
        SelectionHistory.HotkeyGoForward
    End If
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnHistoryForward_Action", Err.Number, Err.Description
End Sub

Public Sub GetHistoryForward_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = SelectionHistory.CanGoForward
End Sub

'===============================================================================
' Profiles
'===============================================================================

Public Sub GetProfilesContent(ByVal control As IRibbonControl, ByRef content)
    On Error GoTo ErrHandler

    Dim index As Long
    Dim menuXml As String
    menuXml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"

    For index = 0 To Profiles.ProfileCount - 1
        menuXml = menuXml & "<button id=""profile" & index & """ label=""" & _
                      EscapeRibbonXml(Profiles.ProfileName(index)) & """ onAction=""OnProfileMenu_Action""/>"
    Next index

    content = menuXml & "</menu>"
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.GetProfilesContent", Err.Number, Err.Description
    content = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui""/>"
End Sub

Public Sub OnProfileMenu_Action(ByVal control As IRibbonControl)
    On Error GoTo ErrHandler

    Dim index As Long
    index = CLng(Mid$(control.id, Len("profile") + 1))
    Profiles.ApplyProfile index
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnProfileMenu_Action", Err.Number, Err.Description
End Sub

Public Sub GetProfile_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.enabled
End Sub

Private Function EscapeRibbonXml(ByVal value As String) As String
    value = Replace(value, "&", "&amp;")
    value = Replace(value, """", "&quot;")
    value = Replace(value, "<", "&lt;")
    value = Replace(value, ">", "&gt;")
    EscapeRibbonXml = value
End Function

'===============================================================================
' Save Profile
'===============================================================================

Public Sub OnSaveProfile_Action(ByVal control As IRibbonControl)
    On Error GoTo ErrHandler
    Dim name As String
    name = InputBox("Profile name:", "Save Profile", "Profile " & (Profiles.ProfileCount + 1))
    If Len(Trim$(name)) > 0 Then
        Profiles.SaveCurrentAsProfile name
        HighlightEngine.ReapplyAllOpenWorkbooks
        InvalidateRibbon
    End If
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnSaveProfile_Action", Err.Number, Err.Description
End Sub

'===============================================================================
' Per-mode colours (row vs column in Crosshair mode)
'===============================================================================

Public Sub OnPerModeColours_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    Settings.PerModeColours = pressed
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnPerModeColours_Action", Err.Number, Err.Description
End Sub

Public Sub GetPerModeColours_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    pressed = Settings.PerModeColours
End Sub

Public Sub GetPerModeColours_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.enabled
End Sub

Public Sub GetRowColourGallery_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.enabled And Settings.PerModeColours
End Sub

Public Sub GetRowColour_SelectedItemID(ByVal control As IRibbonControl, ByRef id)
    If Settings.rowColour = hcCustom Then
        id = ""
    Else
        id = "item" & ColourName(Settings.rowColour)
    End If
End Sub

Public Sub GetColColour_SelectedItemID(ByVal control As IRibbonControl, ByRef id)
    If Settings.colColour = hcCustom Then
        id = ""
    Else
        id = "item" & ColourName(Settings.colColour)
    End If
End Sub

Public Sub OnRowColour_Action(ByVal control As IRibbonControl, ByVal id As String, ByVal index As Integer)
    On Error GoTo ErrHandler
    If Left$(id, 4) = "item" Then
        Settings.rowColour = ColourFromString(Mid$(id, 5))
    End If
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnRowColour_Action", Err.Number, Err.Description
End Sub

Public Sub OnColColour_Action(ByVal control As IRibbonControl, ByVal id As String, ByVal index As Integer)
    On Error GoTo ErrHandler
    If Left$(id, 4) = "item" Then
        Settings.colColour = ColourFromString(Mid$(id, 5))
    End If
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnColColour_Action", Err.Number, Err.Description
End Sub

'===============================================================================
' About
'===============================================================================

Public Sub OnAbout_Action(ByVal control As IRibbonControl)
    On Error Resume Next
    MsgBox "Excel Highlighter" & vbCrLf & _
           "Version " & APP_VERSION & vbCrLf & vbCrLf & _
           "Non-destructive row/column/crosshair highlighting for the active cell." & vbCrLf & _
           "Hotkeys: " & Settings.HotkeyToggle & " toggle | " & _
           Settings.HotkeyHistoryBack & "/" & Settings.HotkeyHistoryFwd & " history back/forward" & vbCrLf & vbCrLf & _
           "https://github.com/Sandeep2062/excel-highlighter-addin", _
           vbInformation, "About"
End Sub
