Attribute VB_Name = "Tests"
Option Explicit

Private mLastTestResults As String

Public Sub RunAllTests()
    Dim results As String
    On Error GoTo ErrHandler

    results = ""

    If AddNumbers(1, 2) = 3 Then
        results = results & "TestAddSimple:PASS" & vbCrLf
    Else
        results = results & "TestAddSimple:FAIL expected 3 got " & CStr(AddNumbers(1, 2)) & vbCrLf
    End If

    mLastTestResults = results
    Exit Sub

ErrHandler:
    mLastTestResults = results & "RunAllTests:EXCEPTION - " & Err.Number & " " & Err.Description
End Sub

Public Function GetTestResults() As String
    If Len(mLastTestResults) = 0 Then
        GetTestResults = "No tests have been executed."
    Else
        GetTestResults = mLastTestResults
    End If
End Function
