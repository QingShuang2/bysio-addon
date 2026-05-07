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

Private Function IsPictureShape(ByVal shp As Object) As Boolean
    On Error GoTo NotPicture
    IsPictureShape = (shp.Type = msoPicture Or shp.Type = msoLinkedPicture)
    Exit Function
NotPicture:
    IsPictureShape = False
End Function

Private Function GetOnlyPictureInSlide(ByVal sld As Object) As Object
    Dim shp As Object
    Dim pic As Object
    Dim picCount As Long

    For Each shp In sld.Shapes
        If IsPictureShape(shp) Then
            picCount = picCount + 1
            Set pic = shp
            If picCount > 1 Then Exit For
        End If
    Next shp

    If picCount = 1 Then
        Set GetOnlyPictureInSlide = pic
    End If
End Function

Private Sub ResizeAndPositionShape(ByVal shp As Object, ByVal targetH As Double, ByVal targetW As Double)
    shp.LockAspectRatio = msoFalse
    shp.Height = targetH
    shp.Width = targetW
    shp.Left = mHorizontal * 72
    shp.Top = mVertical * 72
End Sub

Public Sub RibbonResizeImage_OnAction(ByVal control As Object)
    If mScale = 0 Then InitDefaults

    Dim sel As Object
    Set sel = Application.ActiveWindow.Selection

    ' Scale is based on slide dimensions. All positions are in inches.
    ' 1 inch = 72 points in Office.
    Dim targetH As Double
    Dim targetW As Double
    targetH = SLIDE_H * (mScale / 100) * 72
    targetW = SLIDE_W * (mScale / 100) * 72

    Dim shp As Object
    Dim sld As Object
    Dim onePic As Object
    Dim appliedCount As Long
    Dim skippedCount As Long

    ' ppSelectionShapes = 2
    If sel.Type = 2 Then
        For Each shp In sel.ShapeRange
            If IsPictureShape(shp) Then
                ResizeAndPositionShape shp, targetH, targetW
                appliedCount = appliedCount + 1
            End If
        Next shp

        If appliedCount = 0 Then
            MsgBox "Please select at least one image.", vbExclamation, "Bysio"
        End If
        Exit Sub
    End If

    ' ppSelectionSlides = 1
    If sel.Type = 1 Then
        For Each sld In sel.SlideRange
            Set onePic = GetOnlyPictureInSlide(sld)
            If onePic Is Nothing Then
                skippedCount = skippedCount + 1
            Else
                ResizeAndPositionShape onePic, targetH, targetW
                appliedCount = appliedCount + 1
            End If
        Next sld

        If appliedCount = 0 Then
            MsgBox "No selected slide contains exactly one image.", vbExclamation, "Bysio"
        ElseIf skippedCount > 0 Then
            MsgBox "Updated " & CStr(appliedCount) & " slide(s). Skipped " & CStr(skippedCount) & " slide(s) without exactly one image.", vbInformation, "Bysio"
        End If
        Exit Sub
    End If

    MsgBox "Please select image(s), or select slide(s) that each contain exactly one image.", vbExclamation, "Bysio"
End Sub
