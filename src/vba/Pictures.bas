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

Public Sub ResizeActiveSheetPicturesToPercent(pct As Double)
    Dim ws As Worksheet
    Dim s As Shape
    Dim resizedCount As Long

    On Error Resume Next
    Set ws = Application.ActiveSheet
    If ws Is Nothing Then
        MsgBox "No active sheet to operate on.", vbInformation, APP_TITLE
        Exit Sub
    End If
    On Error GoTo 0

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

    MsgBox resizedCount & " pictures resized to " & pct & "% on sheet '" & ws.Name & "'.", vbInformation, APP_TITLE
End Sub

Public Sub ResizeSelectedPicturesBy(delta As Long)
    Dim sr As Object
    Dim s As Shape
    Dim factor As Double

    On Error Resume Next
    Set sr = Selection.ShapeRange
    If sr Is Nothing Then
        MsgBox "Please select a picture or shape first.", vbInformation, APP_TITLE
        Exit Sub
    End If
    On Error GoTo 0

    factor = 1 + (CDbl(delta) / 100)
    For Each s In sr
        On Error Resume Next
        s.LockAspectRatio = msoTrue
        s.ScaleWidth factor, msoFalse, msoScaleFromTopLeft
    Next s
End Sub

Public Sub ResizeActiveSheetPicturesBy(delta As Long)
    Dim ws As Worksheet
    Dim s As Shape
    Dim resizedCount As Long
    Dim factor As Double

    On Error Resume Next
    Set ws = Application.ActiveSheet
    If ws Is Nothing Then
        MsgBox "No active sheet to operate on.", vbInformation, APP_TITLE
        Exit Sub
    End If
    On Error GoTo 0

    factor = 1 + (CDbl(delta) / 100)
    For Each s In ws.Shapes
        On Error Resume Next
        Select Case s.Type
            Case msoPicture, msoLinkedPicture
                s.LockAspectRatio = msoTrue
                s.ScaleWidth factor, msoFalse, msoScaleFromTopLeft
                resizedCount = resizedCount + 1
        End Select
        On Error GoTo 0
    Next s

    MsgBox resizedCount & " pictures resized by " & delta & "% on sheet '" & ws.Name & "'.", vbInformation, APP_TITLE
End Sub

Public Sub ResizeAllPicturesBy(delta As Long)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim s As Shape
    Dim resizedCount As Long
    Dim factor As Double

    On Error Resume Next
    Set wb = Application.ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "No active workbook to operate on.", vbInformation, APP_TITLE
        Exit Sub
    End If
    On Error GoTo 0

    factor = 1 + (CDbl(delta) / 100)
    For Each ws In wb.Worksheets
        For Each s In ws.Shapes
            On Error Resume Next
            Select Case s.Type
                Case msoPicture, msoLinkedPicture
                    s.LockAspectRatio = msoTrue
                    s.ScaleWidth factor, msoFalse, msoScaleFromTopLeft
                    resizedCount = resizedCount + 1
            End Select
            On Error GoTo 0
        Next s
    Next ws

    MsgBox resizedCount & " pictures resized by " & delta & "% across " & wb.Worksheets.Count & " sheets.", vbInformation, APP_TITLE
End Sub
