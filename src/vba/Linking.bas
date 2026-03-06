Attribute VB_Name = "Linking"
Option Explicit

Public Sub LinkCellsToSheets()
    Dim sel As Range
    Dim c As Range
    Dim wb As Workbook
    Dim targetIndex As Long
    Dim i As Long
    Dim shTarget As Worksheet

    On Error Resume Next
    Set sel = Selection
    If sel Is Nothing Then
        MsgBox "Please select one or more cells first.", vbInformation, APP_TITLE
        Exit Sub
    End If
    On Error GoTo 0

    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook to operate on.", vbInformation, APP_TITLE
        Exit Sub
    End If

    ' Map each row in the selection (top to bottom) to the worksheets after the active sheet
    For i = 1 To sel.Rows.Count
        targetIndex = ActiveSheet.Index + i
        If targetIndex > wb.Worksheets.Count Then
            MsgBox "Not enough sheets after the current sheet to link all selected cells.", vbExclamation, APP_TITLE
            Exit For
        End If

        Set c = sel.Cells(i, 1) ' first column of the selection by row
        Set shTarget = wb.Worksheets(targetIndex)

        On Error Resume Next
        c.Hyperlinks.Delete
        On Error GoTo 0

        wb.Sheets(ActiveSheet.Index).Hyperlinks.Add Anchor:=c, Address:="", SubAddress:="'" & shTarget.Name & "'!A1", TextToDisplay:=shTarget.Name
    Next i
End Sub
