Attribute VB_Name = "Ribbon"
Option Explicit

Private Const APP_TITLE As String = "Bysio Add-in"

Public Sub RibbonAddNumbers_OnAction(control As IRibbonControl)
    Dim firstInput As String
    Dim secondInput As String
    Dim a As Double
    Dim b As Double
    Dim resultValue As Double

    firstInput = InputBox("Enter the first number:", APP_TITLE, "1")
    If Len(firstInput) = 0 Then Exit Sub

    If Not TryParseDouble(firstInput, a) Then
        MsgBox "The first value is not a valid number.", vbExclamation + vbOKOnly, APP_TITLE
        Exit Sub
    End If

    secondInput = InputBox("Enter the second number:", APP_TITLE, "2")
    If Len(secondInput) = 0 Then Exit Sub

    If Not TryParseDouble(secondInput, b) Then
        MsgBox "The second value is not a valid number.", vbExclamation + vbOKOnly, APP_TITLE
        Exit Sub
    End If

    resultValue = AddNumbers(a, b)
    MsgBox "Result: " & Format$(resultValue, "0.############"), vbInformation + vbOKOnly, APP_TITLE
End Sub

Private Function TryParseDouble(ByVal value As String, ByRef parsed As Double) As Boolean
    On Error GoTo ParseFailed
    parsed = CDbl(value)
    TryParseDouble = True
    Exit Function

ParseFailed:
    TryParseDouble = False
End Function
