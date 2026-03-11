Attribute VB_Name = "Ribbon"
Option Explicit

Private mRibbonUI As Object
Private mRibbonFontSelectedIndex As Long
Private mRibbonFontSize As Long

Public Sub RibbonOnLoad(ByVal ribbon As Object)
    Set mRibbonUI = ribbon
    mRibbonFontSize = 11
    Application.StatusBar = "Bysio ribbon loaded."
End Sub

Public Sub RibbonFont_GetSelectedItemIndex(ByVal control As Object, ByRef returnedIndex)
    returnedIndex = mRibbonFontSelectedIndex
End Sub

Public Sub RibbonFont_OnAction(ByVal control As Object, ByVal id As String, ByVal index As Long)
    mRibbonFontSelectedIndex = index
    If Not mRibbonUI Is Nothing Then
        On Error Resume Next
        mRibbonUI.Invalidate
        On Error GoTo 0
    End If
    Dim fontName As String
    Select Case index
        Case 0: fontName = "ＭＳ ゴシック"
        Case 1: fontName = "Meiryo UI"
        Case Else: fontName = ""
    End Select
    Application.StatusBar = "Selected font: " & fontName
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
    Dim fontName As String
    Dim fontSize As Double

    Select Case mRibbonFontSelectedIndex
        Case 0
            fontName = "ＭＳ ゴシック"
            fontSize = mRibbonFontSize
        Case 1
            fontName = "Meiryo UI"
            fontSize = mRibbonFontSize
        Case Else
            PromptAndApplyFont
            Exit Sub
    End Select

    SetActiveSheetFont fontName, fontSize
    MsgBox "Applied font '" & fontName & "' size " & CStr(fontSize) & " to active sheet '" & ActiveSheet.Name & "' in " & ActiveWorkbook.Name, vbInformation
End Sub

Public Sub RibbonZoom100_OnAction(ByVal control As Object)
    ZoomAllSheets100
End Sub

Public Sub RibbonResizePicture_OnAction(ByVal control As Object)
    ResizeAllPicturesToPercent RESIZE_PERCENT
End Sub


Public Sub RibbonSize_GetText(ByVal control As Object, ByRef returnedText)
    returnedText = CStr(mRibbonFontSize)
End Sub

Public Sub RibbonSize_OnChange(ByVal control As Object, ByVal text As String)
    If Len(Trim$(text)) = 0 Then Exit Sub
    If IsNumeric(text) Then
        mRibbonFontSize = CLng(text)
        Application.StatusBar = "Selected font size: " & CStr(mRibbonFontSize)
    Else
        MsgBox "Invalid font size: " & text, vbExclamation
    End If
End Sub



