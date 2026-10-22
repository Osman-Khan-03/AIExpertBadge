Attribute VB_Name = "ModLogging"
Option Explicit

'==============================================================================
' ModLogging - writes timestamped log entries to a "Log" worksheet in this
' workbook and to AutomationLog.txt next to the workbook.
' Logging must never raise an error, so every routine here swallows its own.
'==============================================================================

Public Const LOG_SHEET_NAME As String = "Log"
Private Const LOG_FILE_NAME As String = "AutomationLog.txt"

Public Sub LogInfo(ByVal msg As String)
    WriteLog "INFO", msg
End Sub

Public Sub LogWarning(ByVal msg As String)
    WriteLog "WARNING", msg
End Sub

Public Sub LogError(ByVal msg As String)
    WriteLog "ERROR", msg
End Sub

Public Sub ClearLog()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetOrCreateSheet(ThisWorkbook, LOG_SHEET_NAME)
    If Not ws Is Nothing Then
        ws.Cells.Clear
        WriteLogHeader ws
    End If
End Sub

Private Sub WriteLog(ByVal level As String, ByVal msg As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim stamp As String

    On Error Resume Next

    stamp = Format(Now, "yyyy-mm-dd hh:nn:ss")

    Set ws = GetOrCreateSheet(ThisWorkbook, LOG_SHEET_NAME)
    If Not ws Is Nothing Then
        If IsEmpty(ws.Cells(1, 1).Value) Then WriteLogHeader ws
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(nextRow, 1).Value = stamp
        ws.Cells(nextRow, 2).Value = level
        ws.Cells(nextRow, 3).Value = msg
        If level = "ERROR" Then
            ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, 3)).Font.Color = RGB(192, 0, 0)
        ElseIf level = "WARNING" Then
            ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, 3)).Font.Color = RGB(180, 110, 0)
        End If
    End If

    WriteLogFile stamp & " [" & level & "] " & msg
End Sub

Private Sub WriteLogHeader(ByVal ws As Worksheet)
    ws.Range("A1:C1").Value = Array("Timestamp", "Level", "Message")
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(31, 78, 121)
    ws.Range("A1:C1").Font.Color = RGB(255, 255, 255)
    ws.Columns(1).ColumnWidth = 20
    ws.Columns(2).ColumnWidth = 10
    ws.Columns(3).ColumnWidth = 100
End Sub

Private Sub WriteLogFile(ByVal line As String)
    Dim fileNum As Integer

    On Error Resume Next
    If Len(ThisWorkbook.Path) = 0 Then Exit Sub   ' workbook not saved yet

    fileNum = FreeFile
    Open ThisWorkbook.Path & "\" & LOG_FILE_NAME For Append As #fileNum
    Print #fileNum, line
    Close #fileNum
End Sub
