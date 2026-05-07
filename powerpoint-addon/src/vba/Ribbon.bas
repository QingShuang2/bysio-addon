Attribute VB_Name = "Ribbon"
Option Explicit

Private Const SLIDE_H As Double = 5.85
Private Const SLIDE_W As Double = 9.1

Private mRibbonUI As Object
Private mScale As Double
Private mHorizontal As Double
Private mVertical As Double

Private Sub InitDefaults()
    mScale = 100
    mHorizontal = 3.8
    mVertical = 1.2
End Sub

Public Sub RibbonOnLoad(ByVal ribbon As Object)
    Set mRibbonUI = ribbon
    InitDefaults
End Sub

Public Sub RibbonHorizontal_GetText(ByVal control As Object, ByRef returnedVal)
    If mScale = 0 Then InitDefaults
    returnedVal = CStr(mHorizontal)
End Sub

Public Sub RibbonHorizontal_OnChange(ByVal control As Object, ByVal text As String)
    Dim v As Double
    v = Val(text)
    If v >= 0 Then mHorizontal = v
End Sub

Public Sub RibbonVertical_GetText(ByVal control As Object, ByRef returnedVal)
    If mScale = 0 Then InitDefaults
    returnedVal = CStr(mVertical)
End Sub

Public Sub RibbonVertical_OnChange(ByVal control As Object, ByVal text As String)
    Dim v As Double
    v = Val(text)
    If v >= 0 Then mVertical = v
End Sub

Public Sub RibbonScale_GetText(ByVal control As Object, ByRef returnedVal)
    If mScale = 0 Then InitDefaults
    returnedVal = CStr(mScale)
End Sub

Public Sub RibbonScale_OnChange(ByVal control As Object, ByVal text As String)
    Dim clean As String
    clean = Replace(text, "%", "")
    Dim v As Double
    v = Val(clean)
    If v > 0 Then mScale = v
End Sub

Public Sub RibbonResizeImage_OnAction(ByVal control As Object)
    If mScale = 0 Then InitDefaults

    Dim sel As Object
    Set sel = Application.ActiveWindow.Selection

    ' ppSelectionShapes = 2
    If sel.Type <> 2 Then
        MsgBox "Please select an image first.", vbExclamation, "Bysio"
        Exit Sub
    End If

    ' Scale is based on slide dimensions. All positions are in inches.
    ' 1 inch = 72 points in Office.
    Dim targetH As Double
    Dim targetW As Double
    targetH = SLIDE_H * (mScale / 100) * 72
    targetW = SLIDE_W * (mScale / 100) * 72

    Dim shp As Object
    For Each shp In sel.ShapeRange
        shp.LockAspectRatio = msoFalse
        shp.Height = targetH
        shp.Width = targetW
        shp.Left = mHorizontal * 72
        shp.Top = mVertical * 72
    Next shp
End Sub
