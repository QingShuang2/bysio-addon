Attribute VB_Name = "Formatting"
Option Explicit

Public Sub FormatSelectedNumbers()
    Dim rng As Range
    Dim c As Range

    On Error Resume Next
    Set rng = Selection
    If rng Is Nothing Then
        MsgBox "Please select one or more cells first.", vbInformation, APP_TITLE
        Exit Sub
    End If
    On Error GoTo 0

    For Each c In rng.Cells
        If Not IsError(c.Value) And Len(Trim(CStr(c.Value))) > 0 Then
            If IsNumeric(c.Value) Then
                If Not c.HasFormula Then
                    c.Value = CDbl(c.Value)
                End If
                c.NumberFormat = "General"

                If c.Value = 0 Then
                    c.Interior.Color = RGB(211, 211, 211)
                    c.Font.Color = vbBlack
                ElseIf c.Value > 0 Then
                    c.Interior.Pattern = xlNone
                    c.Font.Color = vbRed
                Else
                    ' Negative numbers: clear special formatting
                    c.Interior.Pattern = xlNone
                    c.Font.Color = vbBlack
                End If
            End If
        End If
    Next c
End Sub
