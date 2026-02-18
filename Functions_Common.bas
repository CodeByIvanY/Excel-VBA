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

Function iFoundRow(cws As Worksheet, iSearch As String, iCol As Long) As Long
    Dim lastRow As Long
    Dim result As Variant
    
    With cws
        lastRow = .Cells(.Rows.Count, iCol).End(xlUp).Row
        result = Application.Match(iSearch, _
                  .Range(.Cells(1, iCol), .Cells(lastRow, iCol)), 0)
    End With
    
    If Not IsError(result) Then
        iFoundRow = result
    Else
        iFoundRow = 0   'Return 0 if not found
    End If

End Function

'Finds the row of a value in a specified column; returns 0 if not found.
Function iFoundRow(cws As Worksheet, iSearch As Variant, iCol As Variant) As Long
    Dim lastRow As Long
    Dim rng As Range
    Dim result As Variant
    Dim colNum As Long

    With cws
        If IsNumeric(iCol) Then
            colNum = CLng(iCol)
        Else
            colNum = .Columns(iCol).Column
        End If
        
        lastRow = .Cells(.Rows.Count, colNum).End(xlUp).Row
        
        Set rng = .Range(.Cells(1, colNum), .Cells(lastRow, colNum))
    End with

    result = Application.Match(iSearch, rng, 0)
    
    If Not IsError(result) Then
        iFoundRow = result
    Else
        iFoundRow = 0
    End If
End Function



