Attribute VB_Name = "Legacy"
Option Explicit

Public Sub Auto_Open()
End Sub

Public Sub Auto_Close()
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

    ' Add Format Numbers legacy button
    Set button = menuBar.Controls.Add(Type:=msoControlButton, Temporary:=True)
    button.Caption = LEGACY_FORMAT_CAPTION
    button.Tag = LEGACY_FORMAT_TAG
    button.Style = msoButtonIconAndCaption
    button.FaceId = 189
    button.OnAction = "RibbonFormatNumbers_LegacyOnAction"

    ' Add Link Cells to Sheets legacy button
    Set button = menuBar.Controls.Add(Type:=msoControlButton, Temporary:=True)
    button.Caption = LEGACY_LINK_CAPTION
    button.Tag = LEGACY_LINK_TAG
    button.Style = msoButtonIconAndCaption
    button.FaceId = 541
    button.OnAction = "RibbonLinkCells_LegacyOnAction"
End Sub

Private Sub RemoveLegacyCommandBarButton()
    On Error Resume Next
    Dim menuBar As CommandBar
    Dim ctrl As CommandBarControl

    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    If menuBar Is Nothing Then Exit Sub

    For Each ctrl In menuBar.Controls
        If ctrl.Tag = LEGACY_APPLY_FONT_TAG Or ctrl.Tag = LEGACY_ZOOM_TAG Or ctrl.Tag = LEGACY_RESIZE_TAG Or ctrl.Tag = LEGACY_FORMAT_TAG Or ctrl.Tag = LEGACY_LINK_TAG Then
            ctrl.Delete
        End If
    Next ctrl
End Sub
