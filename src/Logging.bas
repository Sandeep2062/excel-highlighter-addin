Attribute VB_Name = "Logging"
'===============================================================================
' Module    : Logging
' Purpose   : Very small file-based logger. Deliberately simple - this is a
'             single-user desktop add-in, not a service, so we don't need
'             rotation, levels config UI, etc. Writes to %APPDATA% so it works
'             regardless of where the xlam itself is installed.
'
'             Errors are appended, never shown to the user via MsgBox - the
'             add-in should degrade quietly rather than interrupt someone's
'             spreadsheet work.
'===============================================================================
Option Explicit

Private Const LOG_FILE_NAME As String = "ExcelCrosshairHighlighter.log"
Private Const MAX_LOG_BYTES As Long = 524288   ' 512 KB - trim past this

' Cached so we don't call Environ() on every single log line.
Private mLogPath As String

'-------------------------------------------------------------------------------
' LogPath
' Returns the full path to the log file, creating the folder if required.
'-------------------------------------------------------------------------------
Public Function LogPath() As String

    If Len(mLogPath) = 0 Then
        Dim folder As String
        folder = Environ$("APPDATA") & "\ExcelCrosshairHighlighter"

        If Len(Dir$(folder, vbDirectory)) = 0 Then
            On Error Resume Next
            MkDir folder
            On Error GoTo 0
        End If

        mLogPath = folder & "\" & LOG_FILE_NAME
    End If

    LogPath = mLogPath

End Function

'-------------------------------------------------------------------------------
' LogError
' Description : Records an unexpected error with enough context to diagnose it
'               later without interrupting the user.
' Parameters  : source   - Module.Procedure where the error was trapped
'               errNum   - Err.Number
'               errDesc  - Err.Description
'               extra    - optional free-text context (e.g. sheet name)
' Returns     : None
' Example     : LogError "HighlightEngine.ApplyConditionalFormatting", _
'                         Err.Number, Err.Description, ws.Name
'-------------------------------------------------------------------------------
Public Sub LogError(ByVal source As String, _
                     ByVal errNum As Long, _
                     ByVal errDesc As String, _
                     Optional ByVal extra As String = vbNullString)

    WriteLine "ERROR", source, "#" & errNum & " - " & errDesc & _
              IIf(Len(extra) > 0, " [" & extra & "]", vbNullString)

End Sub

'-------------------------------------------------------------------------------
' LogInfo
' Description : Records a non-error diagnostic message (state transitions,
'               settings changes). Cheap enough to call liberally.
'-------------------------------------------------------------------------------
Public Sub LogInfo(ByVal source As String, ByVal message As String)
    WriteLine "INFO", source, message
End Sub

'-------------------------------------------------------------------------------
' WriteLine
' Internal writer shared by LogError/LogInfo. Swallows its own failures - a
' logging bug must never surface to the user or crash the highlight engine.
'-------------------------------------------------------------------------------
Private Sub WriteLine(ByVal level As String, ByVal source As String, ByVal message As String)

    On Error GoTo CleanFail

    TrimLogIfNeeded

    Dim fnum As Integer
    fnum = FreeFile

    Open LogPath For Append As #fnum
    Print #fnum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & level & vbTab & source & vbTab & message
    Close #fnum

    Exit Sub

CleanFail:
    ' Nothing we can safely do here - don't attempt to log the logging failure.
    On Error Resume Next
    Close #fnum
End Sub

'-------------------------------------------------------------------------------
' TrimLogIfNeeded
' Keeps the log file from growing without bound. Not sophisticated - just
' truncates and starts fresh once the file crosses MAX_LOG_BYTES.
'-------------------------------------------------------------------------------
Private Sub TrimLogIfNeeded()

    On Error Resume Next

    If Len(Dir$(LogPath)) = 0 Then Exit Sub

    If FileLen(LogPath) > MAX_LOG_BYTES Then
        Dim fnum As Integer
        fnum = FreeFile
        Open LogPath For Output As #fnum   ' truncates
        Print #fnum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & "INFO" & vbTab & _
                     "Logging.TrimLogIfNeeded" & vbTab & "Log truncated (exceeded " & MAX_LOG_BYTES & " bytes)"
        Close #fnum
    End If

End Sub
