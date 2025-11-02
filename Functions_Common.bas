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

'PURPOSE: CHECKS IF A WORKSHEET EXISTS IN A GIVEN WORKBOOK
Function WorksheetExists(wsName As String, wb As Workbook) As Boolean
    On Error Resume Next
    WorksheetExists = (wb.Worksheets(wsName).Name = wsName)
End Function

'PURPOSE: TO CONVERTS A NUMERIC COLUMN INDEX INTO ITS CORRESPONDING EXCEL COLUMN LETTER
Function ColLetter(colNum as Long) as String
    Dim iResult as String
    Do While colNum > 0
        colNum = colNum - 1
        iResult = Chr$(65 + (colNum Mod 26)) & iResult
        colNum = colNum \ 26
    Loop
    ColLetter = iResult
End Function
