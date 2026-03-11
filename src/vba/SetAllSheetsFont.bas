Attribute VB_Name = "modSetAllSheetsFont"
Option Explicit

Public Sub SetAllSheetsFont(ByVal fontName As String, ByVal fontSize As Double)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsName As String

    If Application.Workbooks.Count = 0 Then
        MsgBox "No open workbooks.", vbExclamation
        Exit Sub
    End If

    Set wb = ActiveWorkbook
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    For Each ws In wb.Worksheets
        wsName = ws.Name
        If Len(wsName) > 0 Then
            If Left(wsName, 1) = "_" Or Right(wsName, 1) = "_" Then
                ' Skip worksheets whose name starts or ends with an underscore
            Else
                ws.Cells.Font.Name = fontName
                ws.Cells.Font.Size = fontSize
            End If
        End If
    Next ws

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error applying font: " & Err.Description, vbExclamation
End Sub

Public Sub PromptAndApplyFont()
    Const defaultSize As Double = 9
    Dim defaultName As String

    ' Build the Japanese font name with Unicode code points to avoid
    ' encoding/misinterpretation when the .bas file is imported.
    defaultName = ChrW(&HFF2D) & ChrW(&HFF33) & " " & _
                  ChrW(&H30B4) & ChrW(&H30B7) & ChrW(&H30C3) & ChrW(&H30AF)

    SetAllSheetsFont defaultName, defaultSize
    MsgBox "Applied font '" & defaultName & "' size " & CStr(defaultSize) & " to all sheets in " & ActiveWorkbook.Name, vbInformation
End Sub

Public Sub SetActiveSheetFont(ByVal fontName As String, ByVal fontSize As Double)
    Dim ws As Worksheet

    If Application.Workbooks.Count = 0 Then
        MsgBox "No open workbooks.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Set ws = ActiveSheet
    If ws Is Nothing Then
        MsgBox "No active sheet.", vbExclamation
        Exit Sub
    End If

    ws.Cells.Font.Name = fontName
    ws.Cells.Font.Size = fontSize

    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error applying font to active sheet: " & Err.Description, vbExclamation
End Sub

Public Function AddNumbers(a As Double, b As Double) As Double
    AddNumbers = a + b
End Function

Public Sub ZoomAllSheets100()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsName As String
    Dim curSheet As Worksheet

    If Application.Workbooks.Count = 0 Then
        MsgBox "No open workbooks.", vbExclamation
        Exit Sub
    End If

    Set wb = ActiveWorkbook
    On Error Resume Next
    Set curSheet = ActiveSheet
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    For Each ws In wb.Worksheets
        wsName = ws.Name
        If Len(wsName) > 0 Then
            If Left(wsName, 1) = "_" Or Right(wsName, 1) = "_" Then
                ' Skip worksheets starting/ending with underscore
            Else
                ws.Activate
                On Error Resume Next
                ActiveWindow.Zoom = 100
                On Error GoTo ErrHandler
            End If
        End If
    Next ws

    If Not curSheet Is Nothing Then curSheet.Activate
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error setting zoom: " & Err.Description, vbExclamation
End Sub

Public Sub ZoomAllSheetsBy(delta As Long)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsName As String
    Dim curSheet As Worksheet
    Dim curZoom As Long
    Dim newZoom As Long

    If Application.Workbooks.Count = 0 Then
        MsgBox "No open workbooks.", vbExclamation
        Exit Sub
    End If

    Set wb = ActiveWorkbook
    On Error Resume Next
    Set curSheet = ActiveSheet
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    For Each ws In wb.Worksheets
        wsName = ws.Name
        If Len(wsName) > 0 Then
            If Left(wsName, 1) = "_" Or Right(wsName, 1) = "_" Then
                ' Skip worksheets starting/ending with underscore
            Else
                ws.Activate
                On Error Resume Next
                curZoom = ActiveWindow.Zoom
                If curZoom <= 0 Then curZoom = 100
                newZoom = curZoom + delta
                If newZoom < 10 Then newZoom = 10
                If newZoom > 400 Then newZoom = 400
                ActiveWindow.Zoom = newZoom
                On Error GoTo ErrHandler
            End If
        End If
    Next ws

    If Not curSheet Is Nothing Then curSheet.Activate
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error setting zoom: " & Err.Description, vbExclamation
End Sub

Public Sub ZoomActiveSheetBy(delta As Long)
    Dim curZoom As Long
    Dim newZoom As Long

    If Application.Workbooks.Count = 0 Then
        MsgBox "No open workbooks.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    On Error Resume Next
    curZoom = ActiveWindow.Zoom
    If curZoom <= 0 Then curZoom = 100
    newZoom = curZoom + delta
    If newZoom < 10 Then newZoom = 10
    If newZoom > 400 Then newZoom = 400
    ActiveWindow.Zoom = newZoom
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error setting zoom: " & Err.Description, vbExclamation
End Sub

Public Sub ZoomAllSheetsTo(percent As Long)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsName As String
    Dim curSheet As Worksheet

    If Application.Workbooks.Count = 0 Then
        MsgBox "No open workbooks.", vbExclamation
        Exit Sub
    End If

    Set wb = ActiveWorkbook
    On Error Resume Next
    Set curSheet = ActiveSheet
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False

    For Each ws In wb.Worksheets
        wsName = ws.Name
        If Len(wsName) > 0 Then
            If Left(wsName, 1) = "_" Or Right(wsName, 1) = "_" Then
                ' Skip worksheets starting/ending with underscore
            Else
                ws.Activate
                On Error Resume Next
                If percent < 10 Then percent = 10
                If percent > 400 Then percent = 400
                ActiveWindow.Zoom = percent
                On Error GoTo ErrHandler
            End If
        End If
    Next ws

    If Not curSheet Is Nothing Then curSheet.Activate
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error setting zoom: " & Err.Description, vbExclamation
End Sub

Public Sub ZoomActiveSheetTo(percent As Long)
    If Application.Workbooks.Count = 0 Then
        MsgBox "No open workbooks.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    If percent < 10 Then percent = 10
    If percent > 400 Then percent = 400
    On Error Resume Next
    ActiveWindow.Zoom = percent
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "Error setting zoom: " & Err.Description, vbExclamation
End Sub
