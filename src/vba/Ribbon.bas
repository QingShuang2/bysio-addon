Attribute VB_Name = "Ribbon"
Option Explicit

Private mRibbonUI As Object
Private mRibbonFontSelectedIndex As Long
Private mRibbonFontSize As Long
Private mRibbonApplyAllSheets As Boolean
Private mRibbonZoomApplyAllSheets As Boolean

Public Sub RibbonOnLoad(ByVal ribbon As Object)
    Set mRibbonUI = ribbon
    mRibbonFontSize = 11
    mRibbonApplyAllSheets = False
    mRibbonZoomApplyAllSheets = False
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

    If mRibbonApplyAllSheets Then
        SetAllSheetsFont fontName, fontSize
        MsgBox "Applied font '" & fontName & "' size " & CStr(fontSize) & " to all sheets in " & ActiveWorkbook.Name, vbInformation
    Else
        SetActiveSheetFont fontName, fontSize
        MsgBox "Applied font '" & fontName & "' size " & CStr(fontSize) & " to active sheet '" & ActiveSheet.Name & "' in " & ActiveWorkbook.Name, vbInformation
    End If
End Sub

Public Sub RibbonZoom100_OnAction(ByVal control As Object)
    If mRibbonZoomApplyAllSheets Then
        ZoomAllSheets100
    Else
        On Error Resume Next
        ActiveWindow.Zoom = 100
        On Error GoTo 0
    End If
End Sub

Public Sub RibbonResizePicture_OnAction(ByVal control As Object)
    ResizeAllPicturesToPercent RESIZE_PERCENT
End Sub

Public Sub RibbonZoomUp_OnAction(ByVal control As Object)
    If mRibbonZoomApplyAllSheets Then
        ZoomAllSheetsBy 10
    Else
        ZoomActiveSheetBy 10
    End If
End Sub

Public Sub RibbonZoomDown_OnAction(ByVal control As Object)
    If mRibbonZoomApplyAllSheets Then
        ZoomAllSheetsBy -10
    Else
        ZoomActiveSheetBy -10
    End If
End Sub

Public Sub RibbonZoomAllSheets_GetPressed(ByVal control As Object, ByRef returnedPressed)
    returnedPressed = mRibbonZoomApplyAllSheets
End Sub

Public Sub RibbonZoomAllSheets_OnAction(ByVal control As Object, ByVal pressed As Boolean)
    mRibbonZoomApplyAllSheets = pressed
    If Not mRibbonUI Is Nothing Then
        On Error Resume Next
        mRibbonUI.Invalidate
        On Error GoTo 0
    End If
    Application.StatusBar = "Zoom Apply to All Sheets: " & IIf(mRibbonZoomApplyAllSheets, "Yes", "No")
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

Public Sub RibbonAllSheets_GetPressed(ByVal control As Object, ByRef returnedPressed)
    returnedPressed = mRibbonApplyAllSheets
End Sub

Public Sub RibbonAllSheets_OnAction(ByVal control As Object, ByVal pressed As Boolean)
    mRibbonApplyAllSheets = pressed
    If Not mRibbonUI Is Nothing Then
        On Error Resume Next
        mRibbonUI.Invalidate
        On Error GoTo 0
    End If
    Application.StatusBar = "Apply to All Sheets: " & IIf(mRibbonApplyAllSheets, "Yes", "No")
End Sub



