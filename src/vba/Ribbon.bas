Attribute VB_Name = "Ribbon"
Option Explicit

Private Const APP_TITLE As String = "Bysio Add-in"
Private Const LEGACY_POPUP_CAPTION As String = "Bysio Tools"
Private Const LEGACY_BUTTON_CAPTION As String = "Add Numbers"

Public Sub Auto_Open()
    CreateLegacyCommandBarButton
End Sub

Public Sub Auto_Close()
    RemoveLegacyCommandBarButton
End Sub

Public Sub RibbonAddNumbers_OnAction(ByVal control As Object)
    RunAddNumbersPrompt
End Sub

Public Sub RibbonAddNumbers_LegacyOnAction()
    RunAddNumbersPrompt
End Sub

Private Sub RunAddNumbersPrompt()
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

Private Sub CreateLegacyCommandBarButton()
    On Error Resume Next
    RemoveLegacyCommandBarButton
    On Error GoTo 0

    Dim menuBar As CommandBar
    Dim popup As CommandBarControl
    Dim button As CommandBarButton

    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    If menuBar Is Nothing Then Exit Sub

    Set popup = menuBar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    popup.Caption = LEGACY_POPUP_CAPTION

    Set button = popup.Controls.Add(Type:=msoControlButton, Temporary:=True)
    button.Caption = LEGACY_BUTTON_CAPTION
    button.Style = msoButtonIconAndCaption
    button.FaceId = 19
    button.OnAction = "RibbonAddNumbers_LegacyOnAction"
End Sub

Private Sub RemoveLegacyCommandBarButton()
    On Error Resume Next
    Dim menuBar As CommandBar
    Dim ctrl As CommandBarControl

    Set menuBar = Application.CommandBars("Worksheet Menu Bar")
    If menuBar Is Nothing Then Exit Sub

    For Each ctrl In menuBar.Controls
        If ctrl.Caption = LEGACY_POPUP_CAPTION Then
            ctrl.Delete
            Exit For
        End If
    Next ctrl
End Sub

Private Function TryParseDouble(ByVal value As String, ByRef parsed As Double) As Boolean
    On Error GoTo ParseFailed
    parsed = CDbl(value)
    TryParseDouble = True
    Exit Function

ParseFailed:
    TryParseDouble = False
End Function
