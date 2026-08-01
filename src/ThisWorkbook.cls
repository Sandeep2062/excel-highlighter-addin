VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===============================================================================
' Module    : ThisWorkbook
' Purpose   : The add-in's own workbook object. Its only job is to stand up
'             the Application-level event sink (EventApp) when Excel loads
'             the add-in, and tear it down cleanly when the add-in is
'             unloaded. It should never contain feature logic itself.
'
'             gEventApp lives in AddinHost (a standard module) rather than as
'             a local variable here, because a variable declared only in a
'             class module's procedure would go out of scope and be
'             collected as soon as Workbook_Open finishes - which would kill
'             every event handler immediately after the add-in loads.
'===============================================================================
Option Explicit

'-------------------------------------------------------------------------------
' Workbook_Open
' Fires when the xlam is loaded as an add-in (Excel loads add-ins by opening
' them, same as a normal workbook, just without showing a window for it).
'-------------------------------------------------------------------------------
Private Sub Workbook_Open()
    On Error GoTo ErrHandler
    AddinHost.StartUp
    Exit Sub
ErrHandler:
    Logging.LogError "ThisWorkbook.Workbook_Open", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' Workbook_AddinInstall
' Fires once, the first time the user ticks the add-in on in
' Excel Options > Add-ins. Workbook_Open also fires immediately after this,
' so we don't duplicate startup logic here - just log the milestone.
'-------------------------------------------------------------------------------
Private Sub Workbook_AddinInstall()
    On Error Resume Next
    Logging.LogInfo "ThisWorkbook.Workbook_AddinInstall", "Add-in installed, version " & APP_VERSION
End Sub

'-------------------------------------------------------------------------------
' Workbook_AddinUninstall
' Fires when the user unticks the add-in. We proactively strip our
' conditional formatting and names from every currently open workbook so
' nothing is left behind after the add-in disappears.
'-------------------------------------------------------------------------------
Private Sub Workbook_AddinUninstall()
    On Error GoTo ErrHandler
    AddinHost.ShutDown
    Exit Sub
ErrHandler:
    Logging.LogError "ThisWorkbook.Workbook_AddinUninstall", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' Workbook_BeforeClose
' Belt-and-braces cleanup for the (unusual) case where the add-in workbook
' itself is closed directly, e.g. via VBA automation, rather than through the
' normal add-ins UI.
'-------------------------------------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    AddinHost.ShutDown
End Sub

