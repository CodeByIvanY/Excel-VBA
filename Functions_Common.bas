'PURPSE: TURN OFF/ON EXCEL APPLICATION SETTINGS FOR FASTER EXECUTION.
Sub TurnOffApp()
    With Application
        .DisplayAlerts = False
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
       .AskToUpdateLinks = False
    End With
End Sub

Sub TurnOnApp()
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
Function ColLetter(colNum As Long) As String
    Dim c As Byte
    Dim iResult As String
    Do
        c = (colNum - 1) Mod 26
        iResult = Chr$(c + 65) & iResult
        colNum = (colNum - c) \ 26
    Loop While colNum > 0
    ColLetter = iResult
End Function

    'APPLY BORDERS TO ALL SPECIFIED EDGES OF THE RANGE
    Dim edge As Variant
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
        With rng.Borders(edge)
            .LineStyle = xlContinuous
            .TintAndShade = 0
            .ColorIndex = 0 'DEFAULT BLACK COLOR
            .Weight = xlThin
        End With
    Next edge
'iMax is a user-defined VBA function that accepts a variable number of arguments using ParamArray and returns the largest numeric value passed to the function.
Function iMax(ParamArray args() As Variant) as Double
    Dim x As Long
    iMax = args(0)
    For x = 1 To Ubound(args)
        If args(x) > iMax Then
            iMax = args(x)
        End If
    Next X
End Function

Function iMaxArr(arr() as Double) as Double
    Dim x As Byte
    iMax = arr(0)
    For x = 1 To Ubound(arr)
        If arr(x) > iMax Then
            iMax = arr(x)
        End If
    Next X
End Function

