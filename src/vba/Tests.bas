Attribute VB_Name = "Tests"
Option Explicit

Public Sub RunAllTests()
    Dim results As String
    results = ""
    On Error GoTo ErrHandler

    ' Simple functional test
    If AddNumbers(1, 2) = 3 Then
        results = results & "TestAddSimple:PASS" & vbCrLf
    Else
        results = results & "TestAddSimple:FAIL expected 3 got " & CStr(AddNumbers(1, 2)) & vbCrLf
    End If

    ' Write results to a sheet named TestResults
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "TestResults"
    ws.Range("A1").Value = results
    Exit Sub

ErrHandler:
    results = results & "RunAllTests:EXCEPTION - " & Err.Number & " " & Err.Description
    On Error Resume Next
    Dim ws2 As Worksheet
    Set ws2 = ThisWorkbook.Worksheets.Add
    ws2.Name = "TestResults"
    ws2.Range("A1").Value = results
End Sub
