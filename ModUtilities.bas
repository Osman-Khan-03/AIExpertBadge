Attribute VB_Name = "ModUtilities"
Option Explicit

'==============================================================================
' ModUtilities - shared helpers for sheets, workbooks, files and formatting.
'==============================================================================

Private m_PrevCalc As XlCalculation
Private m_FastModeOn As Boolean

'--- Sheets -------------------------------------------------------------------

Public Function SheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    SheetExists = Not ws Is Nothing
End Function

Public Function GetOrCreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    If SheetExists(wb, sheetName) Then
        Set GetOrCreateSheet = wb.Worksheets(sheetName)
    Else
        Set GetOrCreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        GetOrCreateSheet.Name = sheetName
    End If
End Function

Public Function GetLastRow(ByVal ws As Worksheet, Optional ByVal col As Long = 1) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' Last row of the used range (catches rows whose first column is blank).
Public Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    With ws.UsedRange
        GetLastUsedRow = .Row + .Rows.Count - 1
    End With
End Function

Public Function GetLastColumn(ByVal ws As Worksheet, Optional ByVal rowNum As Long = 1) As Long
    GetLastColumn = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
End Function

'--- Workbooks and files ------------------------------------------------------

Public Function FileExists(ByVal path As String) As Boolean
    On Error Resume Next
    FileExists = (Len(path) > 0) And (Len(Dir(path)) > 0)
End Function

' Returns the workbook for the given path, opening it read-only if it is not
' already open. openedHere tells the caller whether it must close it again.
Public Function GetOrOpenWorkbook(ByVal path As String, ByRef openedHere As Boolean) As Workbook
    Dim wb As Workbook

    openedHere = False
    For Each wb In Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            Set GetOrOpenWorkbook = wb
            Exit Function
        End If
    Next wb

    Set GetOrOpenWorkbook = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    openedHere = True
End Function

Public Sub CloseWorkbookQuietly(ByVal wb As Workbook, ByVal doClose As Boolean)
    On Error Resume Next
    If doClose Then
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
    End If
End Sub

' Shows a file picker. Returns "" if the user cancels.
Public Function SelectSourceFile() As String
    Dim result As Variant
    result = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", _
        Title:="Select the source data workbook")
    If VarType(result) = vbBoolean Then
        SelectSourceFile = ""
    Else
        SelectSourceFile = CStr(result)
    End If
End Function

Public Function BuildOutputPath(ByVal sourcePath As String) As String
    Dim folder As String
    folder = Left$(sourcePath, InStrRev(sourcePath, "\") - 1)
    BuildOutputPath = folder & "\Output_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
End Function

'--- Application state --------------------------------------------------------

Public Sub SetFastMode(ByVal enable As Boolean)
    On Error Resume Next
    If enable Then
        If Not m_FastModeOn Then
            m_PrevCalc = Application.Calculation
            m_FastModeOn = True
        End If
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        If m_FastModeOn Then Application.Calculation = m_PrevCalc
        m_FastModeOn = False
    End If
End Sub

'--- Formatting ---------------------------------------------------------------

Public Sub FormatTable(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim header As Range
    Dim body As Range

    Set header = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    Set body = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    With header
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
        .HorizontalAlignment = xlCenter
    End With

    With body.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With

    body.Columns.AutoFit

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    header.AutoFilter

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Select
End Sub
