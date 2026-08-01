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

    ' Register global hotkeys (now configurable via Settings).
    RegisterHotkey
    RegisterHistoryHotkeys

    ' Wire right-click cell context menu.
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
    Settings.enabled = Not Settings.enabled
    HighlightEngine.ReapplyAllOpenWorkbooks
    ' Invalidate the ribbon so toggle button state matches.
    RibbonCallbacks.InvalidateRibbonExternally
    If Settings.enabled Then
        Logging.LogInfo "AddinHost.ToggleHotkeyHandler", "Toggled ON via hotkey"
    Else
        Logging.LogInfo "AddinHost.ToggleHotkeyHandler", "Toggled OFF via hotkey"
    End If
    Exit Sub
ErrHandler:
    Logging.LogError "AddinHost.ToggleHotkeyHandler", Err.Number, Err.Description
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

'-------------------------------------------------------------------------------
' Cell Right-Click Context Menu Integration
'-------------------------------------------------------------------------------
Public Sub AddContextMenu()
    On Error Resume Next
    RemoveContextMenu
    Dim cb As Object
    Set cb = Application.CommandBars("Cell")
    If Not cb Is Nothing Then
        Dim cbut As Object
        Set cbut = cb.Controls.Add(Type:=1, Temporary:=True) ' 1 = msoControlButton
        cbut.Caption = "Toggle Highlighter (Ctrl+Shift+H)"
        cbut.OnAction = "AddinHost.ToggleHotkeyHandler"
        cbut.FaceId = 463
        cbut.BeginGroup = True
    End If
End Sub

Public Sub RemoveContextMenu()
    On Error Resume Next
    Dim cb As Object
    Set cb = Application.CommandBars("Cell")
    If Not cb Is Nothing Then
        Dim ctrl As Object
        For Each ctrl In cb.Controls
            If ctrl.Caption Like "*Highlighter*" Then
                ctrl.Delete
            End If
        Next ctrl
    End If
End Sub
