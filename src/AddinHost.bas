Attribute VB_Name = "AddinHost"
'===============================================================================
' Module    : AddinHost
' Purpose   : Owns the lifetime of the single EventApp instance. Kept
'             separate from ThisWorkbook so the "keep this reference alive"
'             concern is explicit and easy to find, and so RibbonCallbacks
'             can query IsRunning without reaching into ThisWorkbook.
'===============================================================================
Option Explicit

Public gEventApp As EventApp

' Tag used to identify our right-click menu button so we can find and
' remove/refresh it without touching any other menu items.
Private Const CONTEXT_BUTTON_TAG As String = "XLCH_Toggle"

'-------------------------------------------------------------------------------
' StartUp
' Description : Wires up the Application-level event sink. Idempotent - safe
'               to call more than once (e.g. if Workbook_Open somehow fires
'               twice in one session).
'-------------------------------------------------------------------------------
Public Sub StartUp()

    On Error GoTo ErrHandler

    If gEventApp Is Nothing Then
        Set gEventApp = New EventApp
        Set gEventApp.App = Application
        Logging.LogInfo "AddinHost.StartUp", "Event sink attached, version " & APP_VERSION
    End If

    ' Load profiles here, not just in RibbonCallbacks.onLoad: on Excel 2024 /
    ' recent 365 the ribbon is delivered by the COM add-in and its onLoad
    ' forwarding into VBA throws (Application.Run cannot marshal an IRibbonUI
    ' argument), so onLoad never reaches VBA and the Profiles dropdown would
    ' otherwise be empty. Idempotent - safe to also run when onLoad does fire.
    Profiles.Init

    ' Register global hotkeys (now configurable via Settings).
    RegisterHotkey
    RegisterHistoryHotkeys

    ' Add the "Toggle Highlighter" item to the cell right-click menu.
    AddContextMenu

    ' If the user left the add-in enabled last session, make sure the
    ' currently active workbook picks up highlighting immediately rather
    ' than waiting for the next selection change.
    If Settings.enabled Then
        On Error Resume Next
        HighlightEngine.HandleWorkbookOpen ActiveWorkbook
        On Error GoTo ErrHandler
    End If

    Exit Sub

ErrHandler:
    Logging.LogError "AddinHost.StartUp", Err.Number, Err.Description

End Sub

'-------------------------------------------------------------------------------
' AddinFolder
' Description : Folder containing the installed .xlam. ThisWorkbook's
'               GetCustomUI reads customUI14.xml from here (deployed by
'               install.ps1 next to the add-in).
'-------------------------------------------------------------------------------
Public Function AddinFolder() As String
    ' GetCustomUI can fire before the workbook is fully initialised, in which
    ' case ThisWorkbook.Path is empty and the ribbon XML would silently degrade
    ' to the one-button fallback. Fall back to the AddIns collection (which
    ' knows the installed path) so the full ribbon always loads.
    If Len(ThisWorkbook.Path) > 0 Then
        AddinFolder = ThisWorkbook.Path
    Else
        On Error Resume Next
        AddinFolder = Application.AddIns.Item("excel-highlighter.xlam").Path
        If Len(AddinFolder) = 0 Then
            ' Last resort: strip the file name from FullName.
            Dim full As String
            full = ThisWorkbook.FullName
            If InStrRev(full, "\") > 0 Then AddinFolder = Left$(full, InStrRev(full, "\") - 1)
        End If
        On Error GoTo 0
    End If
End Function

'-------------------------------------------------------------------------------
' RegisterHotkey / UnregisterHotkey
' Description : Binds the toggle hotkey (default Ctrl+Shift+H). The key
'               combination is read from Settings so users can customise it
'               via the registry.
'-------------------------------------------------------------------------------
Private Sub RegisterHotkey()
    On Error Resume Next
    Application.OnKey Settings.HotkeyToggle, "AddinHost.ToggleHotkeyHandler"
    If Err.Number = 0 Then
        Logging.LogInfo "AddinHost.RegisterHotkey", "Hotkey " & Settings.HotkeyToggle & " registered"
    End If
    On Error GoTo 0
End Sub

Private Sub UnregisterHotkey()
    On Error Resume Next
    Application.OnKey Settings.HotkeyToggle, ""   ' Clear the binding
    Logging.LogInfo "AddinHost.UnregisterHotkey", "Hotkey " & Settings.HotkeyToggle & " unregistered"
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' RegisterHistoryHotkeys / UnregisterHistoryHotkeys
'-------------------------------------------------------------------------------
Private Sub RegisterHistoryHotkeys()
    On Error Resume Next
    Application.OnKey Settings.HotkeyHistoryBack, "SelectionHistory.HotkeyGoBack"
    Application.OnKey Settings.HotkeyHistoryFwd, "SelectionHistory.HotkeyGoForward"
    Logging.LogInfo "AddinHost.RegisterHistoryHotkeys", "History hotkeys registered"
    On Error GoTo 0
End Sub

Private Sub UnregisterHistoryHotkeys()
    On Error Resume Next
    Application.OnKey Settings.HotkeyHistoryBack, ""
    Application.OnKey Settings.HotkeyHistoryFwd, ""
    Logging.LogInfo "AddinHost.UnregisterHistoryHotkeys", "History hotkeys unregistered"
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' ToggleHotkeyHandler
' Called when the user presses the toggle hotkey. Toggles the enabled state
' and reapplies highlighting accordingly.
'-------------------------------------------------------------------------------
Public Sub ToggleHotkeyHandler()
    On Error GoTo ErrHandler
    ' Route through the same scope-aware toggle the ribbon uses: with
    ' per-workbook scope (default) the hotkey toggles only the active
    ' workbook; with ScopeAll it toggles every open workbook.
    RibbonCallbacks.SetHighlightForActiveWorkbook Not HighlightEngine.ActiveWorkbookHighlightState()
    If HighlightEngine.ActiveWorkbookHighlightState() Then
        Logging.LogInfo "AddinHost.ToggleHotkeyHandler", "Toggled ON via hotkey"
    Else
        Logging.LogInfo "AddinHost.ToggleHotkeyHandler", "Toggled OFF via hotkey"
    End If
    Exit Sub
ErrHandler:
    Logging.LogError "AddinHost.ToggleHotkeyHandler", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' Context menu (cell right-click)
' Adds a native "Toggle Highlighter" item to Excel's Cell command bar so the
' add-in can be switched on/off from the right-click menu, not just the
' ribbon/hotkey. Uses Temporary:=True so the button vanishes automatically
' when Excel closes, and is rebuilt on every StartUp. The checked state is
' refreshed from EventApp.App_SheetBeforeRightClick so it is always current
' when the menu opens.
'-------------------------------------------------------------------------------
Public Sub AddContextMenu()
    On Error GoTo ErrHandler
    Dim cb As CommandBar
    Set cb = Application.CommandBars("Cell")
    If cb Is Nothing Then Exit Sub

    RemoveContextMenuButton cb   ' idempotent - no duplicates across restarts

    Dim btn As CommandBarButton
    Set btn = cb.Controls.Add(Type:=msoControlButton, Temporary:=True)
    btn.Tag = CONTEXT_BUTTON_TAG
    btn.Caption = "Toggle Highlighter"
    btn.BeginGroup = True
    btn.OnAction = "'AddinHost.OnContextToggle'"
    UpdateContextMenuState btn
    Logging.LogInfo "AddinHost.AddContextMenu", "Right-click toggle added"
    Exit Sub
ErrHandler:
    Logging.LogError "AddinHost.AddContextMenu", Err.Number, Err.Description
End Sub

Public Sub RemoveContextMenu()
    On Error Resume Next
    RemoveContextMenuButton Application.CommandBars("Cell")
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' RefreshContextMenuState
' Called by EventApp.App_SheetBeforeRightClick so the checkmark on our menu
' item reflects the real on/off state before the menu is displayed.
'-------------------------------------------------------------------------------
Public Sub RefreshContextMenuState()
    On Error Resume Next
    Dim cb As CommandBar
    Set cb = Application.CommandBars("Cell")
    If cb Is Nothing Then Exit Sub
    Dim btn As CommandBarButton
    Set btn = FindContextMenuButton(cb)
    If Not btn Is Nothing Then UpdateContextMenuState btn
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' OnContextToggle
' Called when the user clicks our right-click menu item. Routes through the
' same scope-aware toggle as the ribbon button and hotkey, so per-workbook
' vs all-workbooks scope rules apply identically.
'-------------------------------------------------------------------------------
Public Sub OnContextToggle()
    On Error GoTo ErrHandler
    RibbonCallbacks.SetHighlightForActiveWorkbook Not HighlightEngine.ActiveWorkbookHighlightState()
    Exit Sub
ErrHandler:
    Logging.LogError "AddinHost.OnContextToggle", Err.Number, Err.Description
End Sub

Private Sub RemoveContextMenuButton(ByVal cb As CommandBar)
    If cb Is Nothing Then Exit Sub
    On Error Resume Next
    Dim i As Long
    For i = cb.Controls.count To 1 Step -1
        If cb.Controls(i).Tag = CONTEXT_BUTTON_TAG Then cb.Controls(i).Delete
    Next i
    On Error GoTo 0
End Sub

Private Function FindContextMenuButton(ByVal cb As CommandBar) As CommandBarButton
    On Error Resume Next
    Dim i As Long
    For i = 1 To cb.Controls.count
        If cb.Controls(i).Tag = CONTEXT_BUTTON_TAG Then
            Set FindContextMenuButton = cb.Controls(i)
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

Private Sub UpdateContextMenuState(ByVal btn As CommandBarButton)
    On Error Resume Next
    If HighlightEngine.ActiveWorkbookHighlightState() Then
        btn.State = msoButtonDown
    Else
        btn.State = msoButtonUp
    End If
    On Error GoTo 0
End Sub

'-------------------------------------------------------------------------------
' ShutDown
' Description : Detaches the event sink and strips our formatting/names from
'               every open workbook. Called on add-in uninstall and as a
'               safety net on workbook close.
'-------------------------------------------------------------------------------
Public Sub ShutDown()

    On Error GoTo ErrHandler

    ' Unregister all hotkeys before tearing down anything else.
    UnregisterHotkey
    UnregisterHistoryHotkeys

    ' Remove our right-click menu item.
    RemoveContextMenu

    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If Utilities.WorkbookIsEligible(wb) Then
            HighlightEngine.HandleWorkbookBeforeClose wb   ' reused: strips CF + names
        End If
    Next wb

    If Not gEventApp Is Nothing Then
        Set gEventApp.App = Nothing
        Set gEventApp = Nothing
        Logging.LogInfo "AddinHost.ShutDown", "Event sink detached"
    End If

    Exit Sub

ErrHandler:
    Logging.LogError "AddinHost.ShutDown", Err.Number, Err.Description

End Sub

'-------------------------------------------------------------------------------
' IsRunning
' Used by RibbonCallbacks.onLoad diagnostics and the About dialog.
'-------------------------------------------------------------------------------
Public Function IsRunning() As Boolean
    IsRunning = Not gEventApp Is Nothing
End Function
