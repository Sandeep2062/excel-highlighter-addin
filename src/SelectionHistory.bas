Attribute VB_Name = "SelectionHistory"
'===============================================================================
' Module    : SelectionHistory
' Purpose   : Tracks the last N cell selections across all workbooks/sheets
'             for back/forward navigation. Each entry stores the workbook
'             name, sheet name, row, and column.
'
'             The history is process-lifetime only (not persisted) since
'             selections are transient. The back/forward hotkeys are
'             Ctrl+Shift+Z and Ctrl+Shift+X.
'===============================================================================
Option Explicit

Private Type HistoryEntry
    wbName  As String
    wsName  As String
    Row     As Long
    Col     As Long
End Type

Private mHistory() As HistoryEntry
Private mCount    As Long
Private mPosition As Long   ' current position in history (-1 = newest)

'-------------------------------------------------------------------------------
' Push
' Description : Records a new selection. If we're not at the newest position,
'               truncates the forward history before adding.
'-------------------------------------------------------------------------------
Public Sub Push(ByVal wbName As String, ByVal wsName As String, _
                ByVal Row As Long, ByVal Col As Long)

    ' Don't record if it's the same as the current entry.
    If mCount > 0 And mPosition >= 0 Then
        If mHistory(mPosition).wbName = wbName And _
           mHistory(mPosition).wsName = wsName And _
           mHistory(mPosition).Row = Row And _
           mHistory(mPosition).Col = Col Then
            Exit Sub
        End If
    End If

    ' Truncate forward history if we navigated back.
    If mPosition < mCount - 1 Then
        mCount = mPosition + 1
    End If

    ' Grow the array if needed.
    If mCount >= SELECTION_HISTORY_SIZE Then
        ' Shift everything left by one.
        Dim i As Long
        For i = 0 To mCount - 2
            mHistory(i) = mHistory(i + 1)
        Next i
        mCount = mCount - 1
        mPosition = mPosition - 1
    End If

    ReDim Preserve mHistory(0 To mCount)
    mHistory(mCount).wbName = wbName
    mHistory(mCount).wsName = wsName
    mHistory(mCount).Row = Row
    mHistory(mCount).Col = Col
    mPosition = mCount
    mCount = mCount + 1

End Sub

'-------------------------------------------------------------------------------
' CanGoBack / CanGoForward
'-------------------------------------------------------------------------------
Public Function CanGoBack() As Boolean
    CanGoBack = (mPosition > 0)
End Function

Public Function CanGoForward() As Boolean
    CanGoForward = (mPosition < mCount - 1)
End Function

'-------------------------------------------------------------------------------
' GoBack / GoForward
' Description : Navigates to the previous/next entry in history. Returns
'               True if navigation occurred.
'-------------------------------------------------------------------------------
Public Function GoBack() As Boolean
    If Not CanGoBack Then
        GoBack = False
        Exit Function
    End If
    mPosition = mPosition - 1
    NavigateToCurrent
    GoBack = True
End Function

Public Function GoForward() As Boolean
    If Not CanGoForward Then
        GoForward = False
        Exit Function
    End If
    mPosition = mPosition + 1
    NavigateToCurrent
    GoForward = True
End Function

'-------------------------------------------------------------------------------
' NavigateToCurrent
' Activates the workbook, sheet, and cell at the current history position.
'-------------------------------------------------------------------------------
Private Sub NavigateToCurrent()

    On Error GoTo ErrHandler

    If mPosition < 0 Or mPosition >= mCount Then Exit Sub

    Dim wb As Workbook
    Set wb = Application.Workbooks(mHistory(mPosition).wbName)
    If wb Is Nothing Then Exit Sub

    wb.Activate

    Dim ws As Worksheet
    Set ws = wb.Worksheets(mHistory(mPosition).wsName)
    If ws Is Nothing Then Exit Sub

    ws.Activate
    ws.Cells(mHistory(mPosition).Row, mHistory(mPosition).Col).Select

    Exit Sub

ErrHandler:
    Logging.LogError "SelectionHistory.NavigateToCurrent", Err.Number, Err.Description

End Sub

'-------------------------------------------------------------------------------
' Clear
'-------------------------------------------------------------------------------
Public Sub Clear()
    mCount = 0
    mPosition = -1
    ReDim mHistory(0 To 0)
End Sub

'-------------------------------------------------------------------------------
' HotkeyGoBack / HotkeyGoForward
' Public entry points for Application.OnKey callbacks.
'-------------------------------------------------------------------------------
Public Sub HotkeyGoBack()
    On Error GoTo ErrHandler
    If GoBack Then
        Logging.LogInfo "SelectionHistory.HotkeyGoBack", "Navigated back"
    End If
    Exit Sub
ErrHandler:
    Logging.LogError "SelectionHistory.HotkeyGoBack", Err.Number, Err.Description
End Sub

Public Sub HotkeyGoForward()
    On Error GoTo ErrHandler
    If GoForward Then
        Logging.LogInfo "SelectionHistory.HotkeyGoForward", "Navigated forward"
    End If
    Exit Sub
ErrHandler:
    Logging.LogError "SelectionHistory.HotkeyGoForward", Err.Number, Err.Description
End Sub
