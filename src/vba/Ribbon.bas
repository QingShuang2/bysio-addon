Attribute VB_Name = "Ribbon"
Option Explicit

Private mRibbonTestInputText As String

Public Sub RibbonTestInput_OnChange(ByVal control As Object, ByVal text As String)
    mRibbonTestInputText = text
End Sub

Public Sub RibbonTestInput_GetText(ByVal control As Object, ByRef returnedVal)
    returnedVal = mRibbonTestInputText
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

Public Sub RibbonResizePicture_OnAction(ByVal control As Object)
    ResizeAllPicturesToPercent RESIZE_PERCENT
End Sub

Public Sub RibbonResizePicture_LegacyOnAction()
    ResizeAllPicturesToPercent RESIZE_PERCENT
End Sub

Public Sub RibbonFormatNumbers_OnAction(ByVal control As Object)
    FormatSelectedNumbers
End Sub

Public Sub RibbonFormatNumbers_LegacyOnAction()
    FormatSelectedNumbers
End Sub

Public Sub RibbonLinkCells_OnAction(ByVal control As Object)
    LinkCellsToSheets
End Sub

Public Sub RibbonLinkCells_LegacyOnAction()
    LinkCellsToSheets
End Sub

