Attribute VB_Name = "ShapeTools"
Option Explicit

Public Sub SetSelectedShapeLineWeight(ByVal points As Single)
    Dim sel As Selection
    Dim i As Long

    If points <= 0 Then
        MsgBox "Line weight must be greater than zero.", vbExclamation, APP_TITLE
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Set sel = ActiveWindow.Selection

    If sel Is Nothing Then
        MsgBox "No active selection.", vbInformation, APP_TITLE
        Exit Sub
    End If

    If sel.Type <> ppSelectionShapes Then
        MsgBox "Please select one or more shapes first.", vbInformation, APP_TITLE
        Exit Sub
    End If

    For i = 1 To sel.ShapeRange.Count
        sel.ShapeRange(i).Line.Visible = msoTrue
        sel.ShapeRange(i).Line.Weight = points
    Next i

    MsgBox "Updated " & CStr(sel.ShapeRange.Count) & " shape(s).", vbInformation, APP_TITLE
    Exit Sub

ErrHandler:
    MsgBox "Error updating shape line weight: " & Err.Description, vbExclamation, APP_TITLE
End Sub
