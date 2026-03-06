Attribute VB_Name = "Ribbon"
Option Explicit

Private Const APP_TITLE As String = "Bysio Add-in"
Private Const LEGACY_BUTTON_CAPTION As String = "Apply Font to All Sheets"
Private Const LEGACY_BUTTON_TAG As String = "BYSIO_APPLY_FONT"

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

Private Sub CreateLegacyCommandBarButton()
    On Error Resume Next
    RemoveLegacyCommandBarButton
    On Error GoTo 0

    Dim menuBar As CommandBar
    Dim button As CommandBarButton

    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    If menuBar Is Nothing Then Exit Sub

    Set button = menuBar.Controls.Add(Type:=msoControlButton, Temporary:=True)
    button.Caption = LEGACY_BUTTON_CAPTION
    button.Tag = LEGACY_BUTTON_TAG
    button.Style = msoButtonIconAndCaption
    button.FaceId = 19
    button.OnAction = "RibbonApplyFont_LegacyOnAction"
End Sub

Private Sub RemoveLegacyCommandBarButton()
    On Error Resume Next
    Dim menuBar As CommandBar
    Dim ctrl As CommandBarControl

    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    If menuBar Is Nothing Then Exit Sub

    For Each ctrl In menuBar.Controls
        If ctrl.Tag = LEGACY_BUTTON_TAG Then
            ctrl.Delete
        End If
    Next ctrl
End Sub

