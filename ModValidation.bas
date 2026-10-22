Attribute VB_Name = "ModValidation"
Option Explicit

'==============================================================================
' ModValidation - checks the source sheet structure and each data row.
' Expected source layout (header in row 1):
'   A: ID (number)   B: Name   C: Date   D: Amount (number)   E: Category
'==============================================================================

Public Const COL_ID As Long = 1
Public Const COL_NAME As Long = 2
Public Const COL_DATE As Long = 3
Public Const COL_AMOUNT As Long = 4
Public Const COL_CATEGORY As Long = 5
Public Const SOURCE_COLUMN_COUNT As Long = 5

Public Function GetExpectedHeaders() As Variant
    GetExpectedHeaders = Array("ID", "Name", "Date", "Amount", "Category")
End Function

' Verifies the sheet has the expected headers and at least one data row.
Public Function ValidateSourceSheet(ByVal ws As Worksheet, ByRef errMsg As String) As Boolean
    Dim expected As Variant
    Dim i As Long
    Dim actual As String

    expected = GetExpectedHeaders()

    For i = 0 To UBound(expected)
        actual = Trim$(CStr(ws.Cells(1, i + 1).Value))
        If StrComp(actual, CStr(expected(i)), vbTextCompare) <> 0 Then
            errMsg = "Header in column " & (i + 1) & " should be '" & expected(i) & _
                     "' but found '" & actual & "'."
            Exit Function
        End If
    Next i

    If GetLastUsedRow(ws) < 2 Then
        errMsg = "The source sheet contains no data rows."
        Exit Function
    End If

    ValidateSourceSheet = True
End Function

' Validates row i of a 2D array read from the source range.
' Returns True if the row is valid; otherwise reason lists every problem found.
Public Function ValidateRow(ByRef data As Variant, ByVal i As Long, ByRef reason As String) As Boolean
    Dim problems As String

    If Not IsNumberValue(data(i, COL_ID)) Then
        problems = AddProblem(problems, "ID is missing or not numeric")
    End If

    If IsBlank(data(i, COL_NAME)) Then
        problems = AddProblem(problems, "Name is missing")
    End If

    If IsError(data(i, COL_DATE)) Then
        problems = AddProblem(problems, "Date is an error value")
    ElseIf IsEmpty(data(i, COL_DATE)) Or Not IsDate(data(i, COL_DATE)) Then
        problems = AddProblem(problems, "Date is missing or invalid")
    End If

    If Not IsNumberValue(data(i, COL_AMOUNT)) Then
        problems = AddProblem(problems, "Amount is missing or not numeric")
    End If

    If IsBlank(data(i, COL_CATEGORY)) Then
        problems = AddProblem(problems, "Category is missing")
    End If

    reason = problems
    ValidateRow = (Len(problems) = 0)
End Function

' True when all source columns of row i are blank.
Public Function IsRowEmpty(ByRef data As Variant, ByVal i As Long) As Boolean
    Dim c As Long
    For c = 1 To SOURCE_COLUMN_COUNT
        If Not IsBlank(data(i, c)) Then Exit Function
    Next c
    IsRowEmpty = True
End Function

Private Function IsBlank(ByVal v As Variant) As Boolean
    If IsError(v) Then
        IsBlank = True
    ElseIf IsEmpty(v) Then
        IsBlank = True
    Else
        IsBlank = (Len(Trim$(CStr(v))) = 0)
    End If
End Function

Private Function IsNumberValue(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbBoolean Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then Exit Function
    End If
    IsNumberValue = IsNumeric(v)
End Function

Private Function AddProblem(ByVal existing As String, ByVal newProblem As String) As String
    If Len(existing) = 0 Then
        AddProblem = newProblem
    Else
        AddProblem = existing & "; " & newProblem
    End If
End Function
