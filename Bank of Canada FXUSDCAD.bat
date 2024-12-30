Option Explicit
'REFERENCES: MICROSOFT SCRIPTING RUNTIME
'CREATED BY: IVAN Y - CI
'PURPOSE: TO FETCHES AND EXTRACT THE AVERAGE FX RATE AND SPOT RATES FROM THE BANK OF CANADA BETWEEN USD AND CAD BASED ON A SPECIFIED DATE RANGE.

Sub FXUSDCAD()
    'URL FOR FETCHING FXUSDCAD DATA
    Const url As String = "https://www.bankofcanada.ca/valet/observations/FXUSDCAD"
    Dim request As Object
    Set request = CreateObject("WinHttp.WinHttpRequest.5.1")
    'SEND GET REQUEST
    request.Open "Get", url, False
    request.Send
    'CHECK FOR SUCCESSFUL RESPONSE
    If request.Status <> 200 Then
        MsgBox request.ResponseText
        Exit Sub
    End If
    
    Dim response As Object
    Set response = JsonConverter.ParseJson(request.ResponseText)
    
    'ACCESS THE DATA
    Dim startDate As Date, endDate As Date
    Dim NumDate As Integer, y As Integer, lastRow As Integer
    Dim AvgFX As Double
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(1)
    'READ START AND END DATES FROM THE WORKSHEET
    With ws
        startDate = .Cells(1, "B").Value
        endDate = .Cells(2, "B").Value
        NumDate = endDate - startDate + 1
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row
        .Range("A5:B" & lastRow + 5).Clear
    End With
    Dim Arr_Data() As Variant
    ReDim Arr_Data(1 To NumDate, 1 To 2) As Variant
    y = 1
    Dim observations As Collection
    Set observations = response("observations")
    
    Dim observation As Object
    'LOOP THROUGH OBSERVATIONS AND FILTER BY DATE
    For Each observation In observations
        If observation("d") >= startDate And observation("d") <= endDate Then
            Arr_Data(y, 1) = observation("d")
            Arr_Data(y, 2) = observation("FXUSDCAD")("v")
            AvgFX = AvgFX + observation("FXUSDCAD")("v")
            y = y + 1
        End If
    Next observation
    'CALCULATE AVERAGE FX RATE
    AvgFX = AvgFX / (y - 1)
    'OUTPUT RESULTS TO THE WORKSHEET
    ws.Range("A5").Resize(UBound(Arr_Data), 2).Value = Arr_Data
    ws.Range("B3").Value = Round(AvgFX, 4)
    MsgBox "Completed", vbInformation
End Sub
