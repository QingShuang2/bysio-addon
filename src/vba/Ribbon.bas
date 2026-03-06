Attribute VB_Name = "Ribbon"
Option Explicit

Private Const APP_TITLE As String = "Bysio Add-in"
Private Const LEGACY_APPLY_FONT_CAPTION As String = "Apply Font to All Sheets"
Private Const LEGACY_APPLY_FONT_TAG As String = "BYSIO_APPLY_FONT"
Private Const LEGACY_ZOOM_CAPTION As String = "Zoom 100% All Sheets"
Private Const LEGACY_ZOOM_TAG As String = "BYSIO_ZOOM_100"
Private Const LEGACY_RESIZE_CAPTION As String = "Resize Picture to 70%"
Private Const LEGACY_RESIZE_TAG As String = "BYSIO_RESIZE_70"

Private Const RESIZE_PERCENT As Double = 70

Public Sub Auto_Open()
    CreateLegacyCommandBarButton
End Sub

Public Sub Auto_Close()
    RemoveLegacyCommandBarButton
End Sub

Public Sub RibbonApplyFont_OnAction(ByVal control As Object)
    PromptAndApplyFont
End Sub

Public Sub RibbonApplyFont_LegacyOnAction()
    PromptAndApplyFont
End Sub

Public Sub RibbonZoom100_OnAction(ByVal control As Object)
    ZoomAllSheets100
End Sub

Public Sub RibbonZoom100_LegacyOnAction()
    ZoomAllSheets100
End Sub

Private Sub CreateLegacyCommandBarButton()
    On Error Resume Next
    RemoveLegacyCommandBarButton
    On Error GoTo 0

    Dim menuBar As CommandBar
    Dim button As CommandBarButton

    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    If menuBar Is Nothing Then Exit Sub

    ' Add Apply Font legacy button
    Set button = menuBar.Controls.Add(Type:=msoControlButton, Temporary:=True)
    button.Caption = LEGACY_APPLY_FONT_CAPTION
    button.Tag = LEGACY_APPLY_FONT_TAG
    button.Style = msoButtonIconAndCaption
    button.FaceId = 19
    button.OnAction = "RibbonApplyFont_LegacyOnAction"

    ' Add Zoom 100% legacy button
    Set button = menuBar.Controls.Add(Type:=msoControlButton, Temporary:=True)
    button.Caption = LEGACY_ZOOM_CAPTION
    button.Tag = LEGACY_ZOOM_TAG
    button.Style = msoButtonIconAndCaption
    button.FaceId = 159
    button.OnAction = "RibbonZoom100_LegacyOnAction"

    ' Add Resize to 70% legacy button
    Set button = menuBar.Controls.Add(Type:=msoControlButton, Temporary:=True)
    button.Caption = LEGACY_RESIZE_CAPTION
    button.Tag = LEGACY_RESIZE_TAG
    button.Style = msoButtonIconAndCaption
    button.FaceId = 260
    button.OnAction = "RibbonResizePicture_LegacyOnAction"
End Sub

Private Sub RemoveLegacyCommandBarButton()
    On Error Resume Next
    Dim menuBar As CommandBar
    Dim ctrl As CommandBarControl

    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    If menuBar Is Nothing Then Exit Sub

    For Each ctrl In menuBar.Controls
        If ctrl.Tag = LEGACY_APPLY_FONT_TAG Or ctrl.Tag = LEGACY_ZOOM_TAG Or ctrl.Tag = LEGACY_RESIZE_TAG Then
            ctrl.Delete
        End If
    Next ctrl
End Sub

Public Sub RibbonResizePicture_OnAction(ByVal control As Object)
    ResizeAllPicturesToPercent RESIZE_PERCENT
End Sub

Public Sub RibbonResizePicture_LegacyOnAction()
    ResizeAllPicturesToPercent RESIZE_PERCENT
End Sub

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

