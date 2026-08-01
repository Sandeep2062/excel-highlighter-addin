VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EventApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
' Class     : EventApp
' Purpose   : Thin Application-level event sink. Every handler here does the
'             minimum possible before delegating to HighlightEngine, so that
'             this class stays easy to scan and HighlightEngine stays the
'             single place that knows how highlighting actually works.
'
'             A single instance of this class is created and kept alive by
'             ThisWorkbook (the add-in's own workbook) for the lifetime of
'             the Excel session - see ThisWorkbook.cls. If no live reference
'             is kept, VBA garbage-collects the object and every event stops
'             firing silently, which is the most common bug in this style of
'             add-in, so don't be tempted to make this local to a procedure.
'===============================================================================
Option Explicit

Public WithEvents App As Application
Attribute App.VB_VarHelpID = -1

'-------------------------------------------------------------------------------
' App_SheetSelectionChange
' Fires for every selection change on every worksheet in every open
' workbook. This is the hot path - see HighlightEngine.HandleSelectionChange
' for the performance notes.
'-------------------------------------------------------------------------------
Private Sub App_SheetSelectionChange(ByVal sh As Object, ByVal target As Range)
    On Error GoTo ErrHandler
    HighlightEngine.HandleSelectionChange sh, target
    Exit Sub
ErrHandler:
    Logging.LogError "EventApp.App_SheetSelectionChange", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' App_WorkbookOpen
'-------------------------------------------------------------------------------
Private Sub App_WorkbookOpen(ByVal wb As Workbook)
    On Error GoTo ErrHandler
    HighlightEngine.HandleWorkbookOpen wb
    Exit Sub
ErrHandler:
    Logging.LogError "EventApp.App_WorkbookOpen", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' App_WorkbookActivate
'-------------------------------------------------------------------------------
Private Sub App_WorkbookActivate(ByVal wb As Workbook)
    On Error GoTo ErrHandler
    HighlightEngine.HandleWorkbookActivate wb
    Exit Sub
ErrHandler:
    Logging.LogError "EventApp.App_WorkbookActivate", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' App_WindowActivate
' NOTE: WindowActivate gives us the workbook via Wb, not the window's sheet
' directly - the sheet-level work happens on the next SelectionChange.
'-------------------------------------------------------------------------------
Private Sub App_WindowActivate(ByVal wb As Workbook, ByVal wn As Window)
    On Error GoTo ErrHandler
    HighlightEngine.HandleWindowActivate wb
    Exit Sub
ErrHandler:
    Logging.LogError "EventApp.App_WindowActivate", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' App_SheetDeactivate
'-------------------------------------------------------------------------------
Private Sub App_SheetDeactivate(ByVal sh As Object)
    On Error GoTo ErrHandler
    HighlightEngine.HandleSheetDeactivate sh
    Exit Sub
ErrHandler:
    Logging.LogError "EventApp.App_SheetDeactivate", Err.Number, Err.Description
End Sub

'-------------------------------------------------------------------------------
' App_WorkbookBeforeClose
' Cleans up names + conditional formatting for the closing workbook so
' nothing of ours survives in the saved file.
'-------------------------------------------------------------------------------
Private Sub App_WorkbookBeforeClose(ByVal wb As Workbook, Cancel As Boolean)
    On Error GoTo ErrHandler
    HighlightEngine.HandleWorkbookBeforeClose wb
    Exit Sub
ErrHandler:
    Logging.LogError "EventApp.App_WorkbookBeforeClose", Err.Number, Err.Description
End Sub
