Attribute VB_Name = "ModDataProcessing"
Option Explicit

'==============================================================================
' ModDataProcessing - Extract, Validate, Transform, Load.
'   Extract   : read the "Data" sheet of the source workbook into an array
'   Validate  : ModValidation checks every row; duplicates are rejected too
'   Transform : trim/normalise text, round amounts, add a Month column
'   Load      : write Output / Rejected / Summary sheets to a new workbook
'==============================================================================

Public Const SOURCE_SHEET_NAME As String = "Data"

Private Const OUT_COLUMN_COUNT As Long = 6
Private Const REJ_COLUMN_COUNT As Long = 7
Private Const STATUS_EVERY As Long = 500

Public Type ProcessResult
    TotalRows As Long
    ValidRows As Long
    RejectedRows As Long
    OutputPath As String
    ErrorMessage As String
End Type

Public Function ProcessData(ByVal sourcePath As String, ByVal outputPath As String, _
                            ByRef result As ProcessResult) As Boolean
    Dim wbSource As Workbook
    Dim wbOut As Workbook
    Dim wsSource As Worksheet
    Dim openedHere As Boolean
    Dim errMsg As String

    Dim data As Variant
    Dim outArr() As Variant
    Dim rejArr() As Variant
    Dim lastRow As Long
    Dim i As Long, c As Long
    Dim validCount As Long, rejCount As Long
    Dim reason As String
    Dim idKey As String
    Dim category As String
    Dim amount As Double
    Dim txDate As Date
    Dim seenIds As Object
    Dim catCount As Object
    Dim catTotal As Object

    On Error GoTo ErrorHandler

    result.TotalRows = 0
    result.ValidRows = 0
    result.RejectedRows = 0
    result.OutputPath = outputPath
    result.ErrorMessage = ""

    LogInfo "Processing started. Source: " & sourcePath

    If Not FileExists(sourcePath) Then
        Err.Raise vbObjectError + 1001, "ProcessData", "Source file not found: " & sourcePath
    End If

    SetFastMode True

    '--- EXTRACT -------------------------------------------------------------
    Set wbSource = GetOrOpenWorkbook(sourcePath, openedHere)

    If Not SheetExists(wbSource, SOURCE_SHEET_NAME) Then
        Err.Raise vbObjectError + 1002, "ProcessData", _
                  "Sheet '" & SOURCE_SHEET_NAME & "' was not found in the source workbook."
    End If
    Set wsSource = wbSource.Worksheets(SOURCE_SHEET_NAME)

    If Not ValidateSourceSheet(wsSource, errMsg) Then
        Err.Raise vbObjectError + 1003, "ProcessData", errMsg
    End If

    lastRow = GetLastUsedRow(wsSource)
    data = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(lastRow, SOURCE_COLUMN_COUNT)).Value
    LogInfo "Extracted " & (lastRow - 1) & " row(s) from sheet '" & SOURCE_SHEET_NAME & "'."

    CloseWorkbookQuietly wbSource, openedHere
    openedHere = False
    Set wbSource = Nothing

    '--- VALIDATE + TRANSFORM ------------------------------------------------
    ReDim outArr(1 To lastRow, 1 To OUT_COLUMN_COUNT)
    ReDim rejArr(1 To lastRow, 1 To REJ_COLUMN_COUNT)

    Set seenIds = CreateObject("Scripting.Dictionary")
    Set catCount = CreateObject("Scripting.Dictionary")
    Set catTotal = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        If Not IsRowEmpty(data, i) Then
            result.TotalRows = result.TotalRows + 1

            If ValidateRow(data, i, reason) Then
                idKey = CStr(CDbl(data(i, COL_ID)))
                If seenIds.Exists(idKey) Then
                    reason = "Duplicate ID (first seen at source row " & seenIds(idKey) & ")"
                Else
                    seenIds.Add idKey, i
                    reason = ""
                End If
            End If

            If Len(reason) = 0 Then
                txDate = Int(CDate(data(i, COL_DATE)))
                amount = Round(CDbl(data(i, COL_AMOUNT)), 2)
                category = UCase$(Trim$(CStr(data(i, COL_CATEGORY))))

                validCount = validCount + 1
                outArr(validCount, 1) = CDbl(data(i, COL_ID))
                outArr(validCount, 2) = StrConv(Trim$(CStr(data(i, COL_NAME))), vbProperCase)
                outArr(validCount, 3) = txDate
                outArr(validCount, 4) = amount
                outArr(validCount, 5) = category
                outArr(validCount, 6) = Format$(txDate, "yyyy-mm")

                If catCount.Exists(category) Then
                    catCount(category) = catCount(category) + 1
                    catTotal(category) = catTotal(category) + amount
                Else
                    catCount.Add category, 1
                    catTotal.Add category, amount
                End If
            Else
                rejCount = rejCount + 1
                rejArr(rejCount, 1) = i
                For c = 1 To SOURCE_COLUMN_COUNT
                    If Not IsError(data(i, c)) Then rejArr(rejCount, c + 1) = data(i, c)
                Next c
                rejArr(rejCount, REJ_COLUMN_COUNT) = reason
                LogWarning "Row " & i & " rejected: " & reason
            End If

            If result.TotalRows Mod STATUS_EVERY = 0 Then
                Application.StatusBar = "Processing row " & i & " of " & lastRow & "..."
            End If
        End If
    Next i

    result.ValidRows = validCount
    result.RejectedRows = rejCount
    LogInfo "Validation complete. Valid: " & validCount & ", Rejected: " & rejCount

    If result.TotalRows = 0 Then
        Err.Raise vbObjectError + 1004, "ProcessData", "The source sheet contains no data rows."
    End If

    '--- LOAD ----------------------------------------------------------------
    Application.StatusBar = "Writing output workbook..."
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    LoadOutputSheet wbOut, outArr, validCount
    LoadRejectedSheet wbOut, rejArr, rejCount
    LoadSummarySheet wbOut, sourcePath, result, catCount, catTotal
    wbOut.Worksheets("Output").Activate

    wbOut.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    LogInfo "Output saved: " & outputPath

    ProcessData = True

CleanUp:
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    CloseWorkbookQuietly wbSource, openedHere
    Application.StatusBar = False
    SetFastMode False
    Exit Function

ErrorHandler:
    result.ErrorMessage = Err.Description
    LogError "ProcessData failed (" & Err.Number & "): " & Err.Description
    ProcessData = False
    Resume CleanUp
End Function

'--- Load helpers -------------------------------------------------------------

Private Sub LoadOutputSheet(ByVal wb As Workbook, ByRef arr() As Variant, ByVal rowCount As Long)
    Dim ws As Worksheet

    Set ws = wb.Worksheets(1)
    ws.Name = "Output"
    ws.Range("A1:F1").Value = Array("ID", "Name", "Date", "Amount", "Category", "Month")

    If rowCount > 0 Then
        ws.Range("C2").Resize(rowCount, 1).NumberFormat = "yyyy-mm-dd"
        ws.Range("D2").Resize(rowCount, 1).NumberFormat = "#,##0.00"
        ws.Range("F2").Resize(rowCount, 1).NumberFormat = "@"
        ws.Range("A2").Resize(rowCount, OUT_COLUMN_COUNT).Value = TrimRows(arr, rowCount, OUT_COLUMN_COUNT)
    End If

    FormatTable ws, rowCount + 1, OUT_COLUMN_COUNT
End Sub

Private Sub LoadRejectedSheet(ByVal wb As Workbook, ByRef arr() As Variant, ByVal rowCount As Long)
    Dim ws As Worksheet

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Rejected"
    ws.Range("A1:G1").Value = Array("Source Row", "ID", "Name", "Date", "Amount", "Category", "Reason")

    If rowCount > 0 Then
        ws.Range("D2").Resize(rowCount, 1).NumberFormat = "yyyy-mm-dd"
        ws.Range("A2").Resize(rowCount, REJ_COLUMN_COUNT).Value = TrimRows(arr, rowCount, REJ_COLUMN_COUNT)
    End If

    FormatTable ws, rowCount + 1, REJ_COLUMN_COUNT
End Sub

' Copies the first rowCount rows of a 1-based 2D array into a right-sized array
' so that unused (empty) rows are not written to the worksheet.
Private Function TrimRows(ByRef arr() As Variant, ByVal rowCount As Long, ByVal colCount As Long) As Variant
    Dim result() As Variant
    Dim r As Long, c As Long

    ReDim result(1 To rowCount, 1 To colCount)
    For r = 1 To rowCount
        For c = 1 To colCount
            result(r, c) = arr(r, c)
        Next c
    Next r
    TrimRows = result
End Function

Private Sub LoadSummarySheet(ByVal wb As Workbook, ByVal sourcePath As String, _
                             ByRef result As ProcessResult, _
                             ByVal catCount As Object, ByVal catTotal As Object)
    Dim ws As Worksheet
    Dim r As Long
    Dim key As Variant

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Summary"

    ws.Range("A1").Value = "Processing Summary"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14

    ws.Range("A3").Value = "Source file"
    ws.Range("B3").Value = sourcePath
    ws.Range("A4").Value = "Processed on"
    ws.Range("B4").Value = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    ws.Range("A5").Value = "Total rows"
    ws.Range("B5").Value = result.TotalRows
    ws.Range("A6").Value = "Valid rows"
    ws.Range("B6").Value = result.ValidRows
    ws.Range("A7").Value = "Rejected rows"
    ws.Range("B7").Value = result.RejectedRows
    ws.Range("A3:A7").Font.Bold = True
    ws.Range("B3:B7").HorizontalAlignment = xlLeft

    ws.Range("A9:C9").Value = Array("Category", "Count", "Total Amount")
    r = 9
    For Each key In catCount.Keys
        r = r + 1
        ws.Cells(r, 1).Value = key
        ws.Cells(r, 2).Value = catCount(key)
        ws.Cells(r, 3).Value = catTotal(key)
    Next key

    If r > 9 Then ws.Range(ws.Cells(10, 3), ws.Cells(r, 3)).NumberFormat = "#,##0.00"
    FormatSummaryTable ws, r
End Sub

' Formats the category table that starts at row 9 on the Summary sheet.
Private Sub FormatSummaryTable(ByVal ws As Worksheet, ByVal lastRow As Long)
    With ws.Range("A9:C9")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 121)
    End With
    With ws.Range(ws.Cells(9, 1), ws.Cells(lastRow, 3)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(191, 191, 191)
    End With
    ws.Columns("A:C").AutoFit
End Sub
