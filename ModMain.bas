Attribute VB_Name = "ModMain"
Option Explicit

'==============================================================================
' ModMain - user-facing entry points.
'   RunAutomation       : pick a source file, process it, show the result
'   RunAutomationOnFile : same, but for a given path (for scheduling/testing)
'   CreateControlPanel  : builds a "Control" sheet with Run / View Log buttons
'   ShowLog             : jumps to the Log sheet
'==============================================================================

Public Sub RunAutomation()
    Dim sourcePath As String

    sourcePath = SelectSourceFile()
    If Len(sourcePath) = 0 Then Exit Sub   ' user cancelled

    RunAutomationOnFile sourcePath
End Sub

Public Sub RunAutomationOnFile(ByVal sourcePath As String)
    Dim result As ProcessResult
    Dim outputPath As String
    Dim msg As String
    Dim answer As VbMsgBoxResult

    outputPath = BuildOutputPath(sourcePath)

    answer = MsgBox("Process this file?" & vbCrLf & vbCrLf & _
                    "Source: " & sourcePath & vbCrLf & _
                    "Output: " & outputPath, _
                    vbOKCancel + vbQuestion, "Data Automation")
    If answer <> vbOK Then Exit Sub

    If ProcessData(sourcePath, outputPath, result) Then
        msg = "Processing complete." & vbCrLf & vbCrLf & _
              "Total rows:    " & result.TotalRows & vbCrLf & _
              "Valid rows:    " & result.ValidRows & vbCrLf & _
              "Rejected rows: " & result.RejectedRows & vbCrLf & vbCrLf & _
              "Saved to: " & result.OutputPath & vbCrLf & vbCrLf & _
              "Open the output file now?"
        If MsgBox(msg, vbYesNo + vbInformation, "Data Automation") = vbYes Then
            Workbooks.Open result.OutputPath
        End If
    Else
        MsgBox "Processing failed:" & vbCrLf & vbCrLf & result.ErrorMessage & vbCrLf & vbCrLf & _
               "See the '" & LOG_SHEET_NAME & "' sheet for details.", _
               vbCritical, "Data Automation"
    End If
End Sub

Public Sub ShowLog()
    If SheetExists(ThisWorkbook, LOG_SHEET_NAME) Then
        ThisWorkbook.Worksheets(LOG_SHEET_NAME).Activate
    Else
        MsgBox "No log entries yet.", vbInformation, "Data Automation"
    End If
End Sub

Public Sub ClearLogWithPrompt()
    If MsgBox("Clear all log entries?", vbYesNo + vbQuestion, "Data Automation") = vbYes Then
        ClearLog
    End If
End Sub

' Builds a simple control sheet with buttons. Run once.
Public Sub CreateControlPanel()
    Dim ws As Worksheet
    Dim btn As Object

    Set ws = GetOrCreateSheet(ThisWorkbook, "Control")
    ws.Cells.Clear

    Do While ws.Buttons.Count > 0
        ws.Buttons(1).Delete
    Loop

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 70

    ws.Range("B2").Value = "Data Automation"
    ws.Range("B2").Font.Size = 18
    ws.Range("B2").Font.Bold = True

    ws.Range("B4").Value = "1. Click 'Run Automation' and choose the source workbook."
    ws.Range("B5").Value = "2. The source needs a '" & SOURCE_SHEET_NAME & "' sheet with headers: ID, Name, Date, Amount, Category."
    ws.Range("B6").Value = "3. A new Output_<timestamp>.xlsx is created next to the source file."
    ws.Range("B7").Value = "4. Rejected rows and reasons are listed on the output's 'Rejected' sheet."

    Set btn = ws.Buttons.Add(ws.Range("B9").Left, ws.Range("B9").Top, 140, 30)
    btn.Caption = "Run Automation"
    btn.OnAction = "RunAutomation"

    Set btn = ws.Buttons.Add(ws.Range("B9").Left + 150, ws.Range("B9").Top, 100, 30)
    btn.Caption = "View Log"
    btn.OnAction = "ShowLog"

    Set btn = ws.Buttons.Add(ws.Range("B9").Left + 260, ws.Range("B9").Top, 100, 30)
    btn.Caption = "Clear Log"
    btn.OnAction = "ClearLogWithPrompt"

    ws.Range("A1").Select
End Sub
