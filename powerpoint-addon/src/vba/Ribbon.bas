Attribute VB_Name = "Ribbon"
Option Explicit

Private mRibbonUI As Object

Public Sub RibbonOnLoad(ByVal ribbon As Object)
    Set mRibbonUI = ribbon
End Sub

Public Sub RibbonOnePlusOne_OnAction(ByVal control As Object)
    MsgBox "1 + 1 = " & CStr(1 + 1), vbInformation, "Bysio"
End Sub
