Attribute VB_Name = "my_xlsx_addins"
' my_xlsx_addins.bas
Option Explicit


Sub CleanBook()
    Dim wb Ad Workbook

    Set wb = ActiveWorkbook

    SpeedUP = True

    Call ActivateCellA1(wb)

    SpeedUp = False

    Call CloseWorkbook(wb, True)
End Sub


Function ActivateCellA1(wb As Workbook)
    Dim ws As Worksheet

    For Each ws in wb.worksheets
        ws.Activate
        ws.Range("A1").Select
        ActiveWindow.Zoom = 100
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    Next

    wb.Worksheets(1).Activate
End Function


Function CloseWorkbook(wb As Workbook, b_shouldSave As Boolean)
    If b_shouldSave = True Then
        Application.DisplayAlerts = False
        wb.Save
        wb.Close
        If Workbooks.Count < 1 Then
            Application.Quit
        Else
            Application.DisplayAlerts = True
        End If
    Else
        Application.DisplayAlerts = False
        wb.Close savechanges:=False
        Application.DisplayAlerts = True
    End If
End Function


Property Let SpeedUp(b_shouldSpeedUp As Boolean)
    Application.EnableEvents = Not b_shouldSpeedUp
    Application.ScreenUpdating = Not b_shouldSpeedUp
    Application.Calculation = IIf(b_shouldSpeedUp, xlCalculationManual, xlCalculationAutomatic)
End Property
