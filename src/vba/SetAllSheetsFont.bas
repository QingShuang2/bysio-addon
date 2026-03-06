Attribute VB_Name = "SetAllSheetsFont"
Option Explicit

Public Sub SetAllSheetsFont(ByVal fontName As String, ByVal fontSize As Double)
    Dim wb As Workbook
    Dim ws As Worksheet

    If Application.Workbooks.Count = 0 Then
        MsgBox "No open workbooks.", vbExclamation
        Exit Sub
    End If

    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    For Each ws In wb.Worksheets
        ws.Cells.Font.Name = fontName
        ws.Cells.Font.Size = fontSize
    Next ws

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error applying font: " & Err.Description, vbExclamation
End Sub

Public Sub PromptAndApplyFont()
    Const defaultName As String = "ＭＳ ゴシック"
    Const defaultSize As Double = 9

    SetAllSheetsFont defaultName, defaultSize
    MsgBox "Applied font '" & defaultName & "' size " & CStr(defaultSize) & " to all sheets in " & ActiveWorkbook.Name, vbInformation
End Sub

Public Function AddNumbers(a As Double, b As Double) As Double
    AddNumbers = a + b
End Function
