'PURPSE: TURN OFF/ON EXCEL APPLICATION SETTINGS FOR FASTER EXECUTION.
Public Sub TurnOffApp()
    With Application
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
       .AskToUpdateLinks = False
    End With
End Sub

Public Sub TurnOnApp()
    With Application
        .DisplayAlerts = True
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
.AskToUpdateLinks = True
    End With
End Sub

Call TurnOffApp
Call TurnOnApp


Function WorksheetExists(wsName As String, wb As Workbook) As Boolean
    On Error Resume Next
    WorksheetExists = (wb.Worksheets(wsName).Name = wsName)
End Function
