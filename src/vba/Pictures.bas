Attribute VB_Name = "Pictures"
Option Explicit

Public Sub ResizeSelectedPicturesToPercent(pct As Double)
    Dim sr As Object
    Dim s As Shape
    On Error Resume Next
    Set sr = Selection.ShapeRange
    If sr Is Nothing Then
        MsgBox "Please select a picture or shape first.", vbInformation, APP_TITLE
        Exit Sub
    End If
    On Error GoTo 0
    For Each s In sr
        On Error Resume Next
        s.LockAspectRatio = msoTrue
        s.ScaleWidth pct / 100, msoTrue, msoScaleFromTopLeft
    Next s
End Sub

Public Sub ResizeAllPicturesToPercent(pct As Double)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim s As Shape
    Dim resizedCount As Long

    On Error Resume Next
    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook to operate on.", vbInformation, APP_TITLE
        Exit Sub
    End If
    On Error GoTo 0

    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            On Error Resume Next
            Select Case s.Type
                Case msoPicture, msoLinkedPicture
                    s.LockAspectRatio = msoTrue
                    s.ScaleWidth pct / 100, msoTrue, msoScaleFromTopLeft
                    resizedCount = resizedCount + 1
            End Select
            On Error GoTo 0
        Next s
    Next ws

    MsgBox resizedCount & " pictures resized to " & pct & "% across " & wb.Worksheets.Count & " sheets.", vbInformation, APP_TITLE
End Sub
