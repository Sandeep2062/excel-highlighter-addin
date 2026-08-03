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
'
' Two paths are used:
'   1. mRibbon.Invalidate - the classic VBA path, when the ribbon was
'      delivered through VBA and onLoad captured a live IRibbonUI.
'   2. COM add-in InvalidateRibbon - on Excel 2024 / recent 365 the ribbon
'      is delivered by the ExcelHighlighter.Ribbon COM add-in, which holds
'      its own IRibbonUI reference. mRibbon here can be Nothing or stale
'      (the Application.Run bridge does not reliably hand the interface
'      through), so invalidating via the COM object is the reliable refresh.
'      Without it, toggling from the context menu or hotkey changes the
'      state but the ribbon keeps showing the old pressed/label state.
'-------------------------------------------------------------------------------
Private Sub InvalidateRibbon()
    On Error Resume Next
    Dim viaVba As Boolean
    If Not mRibbon Is Nothing Then
        mRibbon.Invalidate
        viaVba = True
    End If
    ' Refresh through the COM add-in that owns the ribbon on Excel 2024.
    ' Two channels into it are empirically broken on this build:
    '   1. Application.COMAddIns("...").Object returns Nothing.
    '   2. Application.Run cannot marshal an IRibbonUI argument (so onLoad
    '      never hands VBA a live mRibbon).
    ' The working route: the COM add-in keeps the ribbon in a STATIC field
    ' (set by the connected instance's OnLoad), so a throwaway instance
    ' created here shares it and invalidates the real ribbon.
    Dim comRibbon As Object
    Set comRibbon = CreateObject("ExcelHighlighter.Ribbon")
    If Not comRibbon Is Nothing Then comRibbon.InvalidateRibbon
    ' Only log when neither route worked - the working case fires on every
    ' ribbon action and would spam the log otherwise.
    If Not viaVba And comRibbon Is Nothing Then
        Logging.LogInfo "RibbonCallbacks.InvalidateRibbon", _
            "No refresh route available (mRibbon invalid, CreateObject returned Nothing)"
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' RibbonAttached
' Diagnostics: whether onLoad captured a live IRibbonUI reference.
'-------------------------------------------------------------------------------
Public Function RibbonAttached() As Boolean
    RibbonAttached = Not mRibbon Is Nothing
End Function

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
    SetHighlightForActiveWorkbook pressed
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnToggle_Action", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' SetHighlightForActiveWorkbook
' The single place that decides what "toggling the highlight" means.
'   - ScopeAll=False (default, per-workbook): the toggle affects ONLY the
'     workbook that is active when it is pressed. Other open workbooks are
'     not touched. Turning the last enabled workbook off turns the master
'     switch off too (so the engine stops running entirely).
'   - ScopeAll=True ("All Workbooks"): the toggle affects every open
'     workbook at once, and turning it off from any workbook turns it off
'     for all of them.
' Used by the ribbon button AND the Ctrl+Shift+H hotkey (AddinHost).
'-------------------------------------------------------------------------------
Public Sub SetHighlightForActiveWorkbook(ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    If Settings.ScopeAll Then
        Settings.enabled = pressed
        HighlightEngine.ReapplyAllOpenWorkbooks
    Else
        Dim wb As Workbook
        Set wb = ActiveWorkbook
        If wb Is Nothing Then Exit Sub
        If pressed Then
            Settings.enabled = True
            HighlightEngine.SetWorkbookEnabled wb, True
        Else
            HighlightEngine.SetWorkbookEnabled wb, False
            ' If no other open workbook is enabled, stop the engine entirely.
            If Not AnyOpenWorkbookEnabled Then Settings.enabled = False
        End If
        HighlightEngine.ReapplyAllOpenWorkbooks
    End If
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.SetHighlightForActiveWorkbook", Err.Number, Err.Description
End Sub

Private Function AnyOpenWorkbookEnabled() As Boolean
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If HighlightEngine.WorkbookEnabled(wb) Then
            AnyOpenWorkbookEnabled = True
            Exit Function
        End If
    Next wb
End Function

Public Sub GetToggle_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    On Error Resume Next
    pressed = HighlightEngine.ActiveWorkbookHighlightState()
End Sub

Public Sub GetToggle_Label(ByVal control As IRibbonControl, ByRef label)
    On Error Resume Next
    If HighlightEngine.ActiveWorkbookHighlightState() Then
        label = "Highlight: On"
    Else
        label = "Highlight: Off"
    End If
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
' Highlight scope (per-workbook vs all open workbooks)
'===============================================================================

Public Sub OnScopeAll_Action(ByVal control As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo ErrHandler
    Settings.ScopeAll = pressed
    ' Switching to "all workbooks" also turns the master on everywhere.
    ' Switching back to per-workbook keeps whatever workbooks are currently
    ' enabled in their per-workbook state.
    If pressed Then
        Settings.enabled = True
        Dim wb As Workbook
        For Each wb In Application.Workbooks
            HighlightEngine.SetWorkbookEnabled wb, True
        Next wb
    End If
    HighlightEngine.ReapplyAllOpenWorkbooks
    InvalidateRibbon
    Exit Sub
ErrHandler:
    Logging.LogError "RibbonCallbacks.OnScopeAll_Action", Err.Number, Err.Description
End Sub

Public Sub GetScopeAll_Pressed(ByVal control As IRibbonControl, ByRef pressed)
    pressed = Settings.ScopeAll
End Sub

Public Sub GetScopeAll_Label(ByVal control As IRibbonControl, ByRef label)
    If Settings.ScopeAll Then
        label = "All Workbooks"
    Else
        label = "This Workbook"
    End If
End Sub

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
        ' Also sweep the legacy pre-2.1.4 worksheet-scoped name.
        Utilities.SafeDeleteName ws.Parent, NAME_LEGACY_SHEET_EXCLUDED
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
        Utilities.DeleteLegacyNames wb
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

'===============================================================================
' COM-bridge *Value functions
' On Excel 2024 / recent Microsoft 365 builds the VBA ribbon delivery path
' (customUI part and ThisWorkbook IRibbonExtensibility) is silently never
' queried, so the ribbon is delivered by the ExcelHighlighter.Ribbon COM
' add-in instead. Every ribbon callback resolves to a method on that COM
' object, which delegates here via Application.Run. Application.Run passes
' arguments BY VALUE, so the Sub + ByRef pattern the VBA-only path uses
' cannot return values through it - these thin Functions call the existing
' Subs and return the ByRef result as the function value. The Subs above
' remain the single source of logic.
'===============================================================================

Public Function GetToggle_PressedValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetToggle_Pressed(control, v)
    GetToggle_PressedValue = CBool(v)
End Function

Public Function GetToggle_LabelValue(ByVal control As IRibbonControl) As String
    Dim v
    Call GetToggle_Label(control, v)
    GetToggle_LabelValue = CStr(v)
End Function

Public Function GetMode_PressedValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetMode_Pressed(control, v)
    GetMode_PressedValue = CBool(v)
End Function

Public Function GetMode_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetMode_Enabled(control, v)
    GetMode_EnabledValue = CBool(v)
End Function

Public Function GetGallery_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetGallery_Enabled(control, v)
    GetGallery_EnabledValue = CBool(v)
End Function

Public Function GetGallery_SelectedItemIDValue(ByVal control As IRibbonControl) As String
    Dim v
    Call GetGallery_SelectedItemID(control, v)
    GetGallery_SelectedItemIDValue = CStr(v)
End Function

Public Function GetGallery_ItemCountValue(ByVal control As IRibbonControl) As Long
    Dim v
    Call GetGallery_ItemCount(control, v)
    GetGallery_ItemCountValue = CLng(v)
End Function

Public Function GetGallery_ItemIDValue(ByVal control As IRibbonControl, ByVal index As Integer) As String
    Dim v
    Call GetGallery_ItemID(control, index, v)
    GetGallery_ItemIDValue = CStr(v)
End Function

Public Function GetGallery_ItemLabelValue(ByVal control As IRibbonControl, ByVal index As Integer) As String
    Dim v
    Call GetGallery_ItemLabel(control, index, v)
    GetGallery_ItemLabelValue = CStr(v)
End Function

Public Function GetGallery_ItemScreentipValue(ByVal control As IRibbonControl, ByVal index As Integer) As String
    Dim v
    Call GetGallery_ItemScreentip(control, index, v)
    GetGallery_ItemScreentipValue = CStr(v)
End Function

Public Function GetGallery_ItemImageValue(ByVal control As IRibbonControl, ByVal index As Integer) As Object
    Dim v As Object
    Call GetGallery_ItemImage(control, index, v)
    Set GetGallery_ItemImageValue = v
End Function

'-------------------------------------------------------------------------------
' GetGallery_ItemRGBValue
' COM-bridge helper: the C# ribbon add-in generates the gallery swatch
' bitmaps itself (System.Drawing - no GDI/OleCreatePictureIndirect risk on
' 64-bit Excel), so it only needs the RGB value of each gallery item from
' VBA, not a picture. Returns the preset or recent-colour RGB for index.
'-------------------------------------------------------------------------------
Public Function GetGallery_ItemRGBValue(ByVal control As IRibbonControl, ByVal index As Integer) As Long
    If index < PRESET_COLOUR_COUNT Then
        GetGallery_ItemRGBValue = PresetColourRGB(index)
    Else
        Dim recentIndex As Long
        recentIndex = index - PRESET_COLOUR_COUNT
        If recentIndex >= 0 And recentIndex < Settings.RecentColourCount Then
            GetGallery_ItemRGBValue = Settings.RecentColour(recentIndex)
        Else
            GetGallery_ItemRGBValue = RGB_YELLOW
        End If
    End If
End Function

'-------------------------------------------------------------------------------
' GetSwatchRGBValue
' COM-bridge helper: RGB for the custom-colour button swatch (the effective
' highlight colour, after any dark-mode tint).
'-------------------------------------------------------------------------------
Public Function GetSwatchRGBValue(ByVal control As IRibbonControl) As Long
    GetSwatchRGBValue = Settings.EffectiveRGB
End Function

Public Function GetScopeAll_PressedValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetScopeAll_Pressed(control, v)
    GetScopeAll_PressedValue = CBool(v)
End Function

Public Function GetScopeAll_LabelValue(ByVal control As IRibbonControl) As String
    Dim v
    Call GetScopeAll_Label(control, v)
    GetScopeAll_LabelValue = CStr(v)
End Function

Public Function GetPerModeColours_PressedValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetPerModeColours_Pressed(control, v)
    GetPerModeColours_PressedValue = CBool(v)
End Function

Public Function GetPerModeColours_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetPerModeColours_Enabled(control, v)
    GetPerModeColours_EnabledValue = CBool(v)
End Function

Public Function GetRowColourGallery_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetRowColourGallery_Enabled(control, v)
    GetRowColourGallery_EnabledValue = CBool(v)
End Function

Public Function GetRowColour_SelectedItemIDValue(ByVal control As IRibbonControl) As String
    Dim v
    Call GetRowColour_SelectedItemID(control, v)
    GetRowColour_SelectedItemIDValue = CStr(v)
End Function

Public Function GetColColour_SelectedItemIDValue(ByVal control As IRibbonControl) As String
    Dim v
    Call GetColColour_SelectedItemID(control, v)
    GetColColour_SelectedItemIDValue = CStr(v)
End Function

Public Function GetExclude_PressedValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetExclude_Pressed(control, v)
    GetExclude_PressedValue = CBool(v)
End Function

Public Function GetExclude_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetExclude_Enabled(control, v)
    GetExclude_EnabledValue = CBool(v)
End Function

Public Function GetExclude_LabelValue(ByVal control As IRibbonControl) As String
    Dim v
    Call GetExclude_Label(control, v)
    GetExclude_LabelValue = CStr(v)
End Function

Public Function GetExcludeSheet_PressedValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetExcludeSheet_Pressed(control, v)
    GetExcludeSheet_PressedValue = CBool(v)
End Function

Public Function GetExcludeSheet_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetExcludeSheet_Enabled(control, v)
    GetExcludeSheet_EnabledValue = CBool(v)
End Function

Public Function GetExcludeSheet_LabelValue(ByVal control As IRibbonControl) As String
    Dim v
    Call GetExcludeSheet_Label(control, v)
    GetExcludeSheet_LabelValue = CStr(v)
End Function

Public Function GetAllowProtected_PressedValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetAllowProtected_Pressed(control, v)
    GetAllowProtected_PressedValue = CBool(v)
End Function

Public Function GetAllowProtected_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetAllowProtected_Enabled(control, v)
    GetAllowProtected_EnabledValue = CBool(v)
End Function

Public Function GetDarkMode_PressedValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetDarkMode_Pressed(control, v)
    GetDarkMode_PressedValue = CBool(v)
End Function

Public Function GetDarkMode_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetDarkMode_Enabled(control, v)
    GetDarkMode_EnabledValue = CBool(v)
End Function

Public Function GetStyle_PressedValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetStyle_Pressed(control, v)
    GetStyle_PressedValue = CBool(v)
End Function

Public Function GetStyle_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetStyle_Enabled(control, v)
    GetStyle_EnabledValue = CBool(v)
End Function

Public Function GetIntersection_PressedValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetIntersection_Pressed(control, v)
    GetIntersection_PressedValue = CBool(v)
End Function

Public Function GetIntersection_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetIntersection_Enabled(control, v)
    GetIntersection_EnabledValue = CBool(v)
End Function

Public Function GetAnimated_PressedValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetAnimated_Pressed(control, v)
    GetAnimated_PressedValue = CBool(v)
End Function

Public Function GetAnimated_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetAnimated_Enabled(control, v)
    GetAnimated_EnabledValue = CBool(v)
End Function

Public Function GetHistoryBack_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetHistoryBack_Enabled(control, v)
    GetHistoryBack_EnabledValue = CBool(v)
End Function

Public Function GetHistoryForward_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetHistoryForward_Enabled(control, v)
    GetHistoryForward_EnabledValue = CBool(v)
End Function

Public Function GetProfile_EnabledValue(ByVal control As IRibbonControl) As Boolean
    Dim v
    Call GetProfile_Enabled(control, v)
    GetProfile_EnabledValue = CBool(v)
End Function

Public Function GetProfilesContentValue(ByVal control As IRibbonControl) As String
    Dim v
    Call GetProfilesContent(control, v)
    GetProfilesContentValue = CStr(v)
End Function
