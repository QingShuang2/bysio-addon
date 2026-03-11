Attribute VB_Name = "Ribbon"
Option Explicit

Private mRibbonUI As Object
Private mRibbonTestInputText As String

Public Sub RibbonOnLoad(ByVal ribbon As Object)
    Set mRibbonUI = ribbon
    Application.StatusBar = "Bysio ribbon loaded."
End Sub

Public Sub RibbonTestInput_OnChange(ByVal control As Object, ByVal text As String)
    mRibbonTestInputText = text
End Sub

Public Sub RibbonTestInput_GetText(ByVal control As Object, ByRef returnedVal)
    returnedVal = mRibbonTestInputText
End Sub

Public Sub RibbonCustomTabTest_OnAction(ByVal control As Object)
    If Not mRibbonUI Is Nothing Then
        On Error Resume Next
        mRibbonUI.Invalidate
        On Error GoTo 0
    End If
    Application.StatusBar = "Bysio custom ribbon tab loaded."
End Sub

Public Sub RibbonApplyFont_OnAction(ByVal control As Object)
    PromptAndApplyFont
End Sub

Public Sub RibbonZoom100_OnAction(ByVal control As Object)
    ZoomAllSheets100
End Sub

Public Sub RibbonResizePicture_OnAction(ByVal control As Object)
    ResizeAllPicturesToPercent RESIZE_PERCENT
End Sub

Public Sub RibbonFormatNumbers_OnAction(ByVal control As Object)
    FormatSelectedNumbers
End Sub


