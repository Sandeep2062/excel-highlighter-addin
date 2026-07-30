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

' The number of preset colours in the gallery.
Private Const PRESET_COLOUR_COUNT As Long = 7

Public Sub OnGallery_Action(ByVal control As IRibbonControl, ByVal id As String, ByVal index As Integer)

    On Error GoTo ErrHandler

    If Left$(id, 4) = "item" Then
        ' Preset colour selected.
        Settings.Colour = ColourFromString(Mid$(id, 5))
    ElseIf Left$(id, 6) = "recent" Then
        ' Recent colour selected. Apply it as custom.
        Dim recentIndex As Long
        recentIndex = CLng(Mid$(id, 7))
        Dim recentRGB As Long
        Dim colours As Long
        colours = Settings.RecentColours
        If recentIndex >= 0 And recentIndex < Settings.RecentColourCount Then
            Settings.CustomRGB = Settings.RecentColours(recentIndex)
            Settings.Colour = hcCustom
        End If
    End If

    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon

    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnGallery_Action", Err.Number, Err.Description
End Sub

Public Sub GetGallery_SelectedItemID(ByVal control As IRibbonControl, ByRef id)
    If Settings.Colour = hcCustom Then
        ' Custom colour - select Nothing (no preset), or highlight the
        ' first recent colour that matches.
        id = ""
    Else
        id = "item" & ColourName(Settings.Colour)
    End If
End Sub

Public Sub GetGallery_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.Enabled
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
' For presets: loads the PNG from disk. For recent colours: generates a
' coloured bitmap in memory since we don't have a PNG for arbitrary RGB.
'-------------------------------------------------------------------------------
Public Sub GetGallery_ItemImage(ByVal control As IRibbonControl, ByVal index As Integer, ByRef image)

    On Error GoTo ErrHandler

    If index < PRESET_COLOUR_COUNT Then
        ' Load preset PNG from disk.
        Dim path As String
        path = ImagesFolder() & LCase$(PresetColourName(index)) & ".png"
        If Len(Dir$(path)) > 0 Then
            Set image = LoadPicture(path)
        End If
    Else
        ' Generate a coloured image for recent colours.
        Dim recentIndex As Long
        recentIndex = index - PRESET_COLOUR_COUNT
        If recentIndex >= 0 And recentIndex < Settings.RecentColourCount Then
            Dim rgbVal As Long
            rgbVal = Settings.RecentColours(recentIndex)
            Set image = GenerateColourSwatch(rgbVal)
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

'-------------------------------------------------------------------------------
' GetSwatchImage
' Shared getImage handler for the custom colour button (not the gallery items
' themselves, which use getItemImage now).
'-------------------------------------------------------------------------------
Public Sub GetSwatchImage(ByVal control As IRibbonControl, ByRef image)

    On Error GoTo ErrHandler

    Dim path As String
    If control.ID = CTRL_CUSTOM_COLOUR Then
        path = ImagesFolder() & "custom.png"
    Else
        path = ImagesFolder() & LCase$(Mid$(control.ID, Len("item") + 1)) & ".png"
    End If

    If Len(Dir$(path)) > 0 Then
        Set image = LoadPicture(path)
    End If

    Exit Sub

ErrHandler:
    Logging.LogError "RibbonCallbacks.GetSwatchImage", Err.Number, Err.Description, control.ID

End Sub

Private Function ImagesFolder() As String
    ImagesFolder = ThisWorkbook.Path & Application.PathSeparator & "images" & Application.PathSeparator
End Function

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
        Settings.Colour = hcCustom
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
        Logging.LogInfo "RibbonCallbacks.OnExcludeSheet_Action", "Excluded sheet " & ws.Name
    Else
        ' Re-enable highlighting on this sheet.
        If Settings.Enabled And Settings.Mode <> hmNone Then
            HighlightEngine.RebuildConditionalFormatting ws
        End If
        Logging.LogInfo "RibbonCallbacks.OnExcludeSheet_Action", "Un-excluded sheet " & ws.Name
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
        Logging.LogInfo "RibbonCallbacks.OnExclude_Action", "Excluded " & wb.Name
    Else
        ' Re-enable highlighting on this workbook.
        If Settings.Enabled And Settings.Mode <> hmNone Then
            HighlightEngine.EnsureNamesExist wb
            On Error Resume Next
            If TypeOf wb.ActiveSheet Is Worksheet Then
                HighlightEngine.RebuildConditionalFormatting wb.ActiveSheet
            End If
            On Error GoTo ErrHandler
        End If
        Logging.LogInfo "RibbonCallbacks.OnExclude_Action", "Un-excluded " & wb.Name
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
    requestedStyle = StyleForControlId(control.ID)
    Settings.HighlightStyle = requestedStyle
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnStyle_Action", Err.Number, Err.Description
End Sub

Public Sub GetStyle_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    pressed = (Settings.HighlightStyle = StyleForControlId(control.ID))
End Sub

Public Sub GetStyle_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.Enabled
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
    enabled = Settings.Enabled
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
    enabled = Settings.Enabled
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
    enabled = Settings.Enabled
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
    enabled = Settings.Enabled
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
                  EscapeRibbonXml(Profiles.ProfileName(index)) & """ onAction=""RibbonCallbacks.OnProfileMenu_Action""/>"
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
    index = CLng(Mid$(control.ID, Len("profile") + 1))
    Profiles.ApplyProfile index
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnProfileMenu_Action", Err.Number, Err.Description
End Sub

Public Sub GetProfile_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.Enabled
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
    enabled = Settings.Enabled
End Sub

Public Sub GetRowColourGallery_Enabled(ByVal control As IRibbonControl, ByRef enabled)
    enabled = Settings.Enabled And Settings.PerModeColours
End Sub

Public Sub GetRowColour_SelectedItemID(ByVal control As IRibbonControl, ByRef id)
    If Settings.RowColour = hcCustom Then
        id = ""
    Else
        id = "item" & ColourName(Settings.RowColour)
    End If
End Sub

Public Sub GetColColour_SelectedItemID(ByVal control As IRibbonControl, ByRef id)
    If Settings.ColColour = hcCustom Then
        id = ""
    Else
        id = "item" & ColourName(Settings.ColColour)
    End If
End Sub

Public Sub OnRowColour_Action(ByVal control As IRibbonControl, ByVal id As String, ByVal index As Integer)
    On Error GoTo ErrHandler
    If Left$(id, 4) = "item" Then
        Settings.RowColour = ColourFromString(Mid$(id, 5))
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
        Settings.ColColour = ColourFromString(Mid$(id, 5))
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
    MsgBox "Excel-Highlighter" & vbCrLf & _
           "Version " & APP_VERSION & vbCrLf & vbCrLf & _
           "Non-destructive row/column/crosshair highlighting for the active cell." & vbCrLf & _
           "Hotkeys: " & Settings.HotkeyToggle & " toggle | " & _
           Settings.HotkeyHistoryBack & "/" & Settings.HotkeyHistoryFwd & " history back/forward" & vbCrLf & vbCrLf & _
           "https://github.com/Sandeep2062/excel-highlighter-addin", _
           vbInformation, "About"
End Sub
