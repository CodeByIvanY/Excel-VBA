Private Enum RowWks
    Tasksize = 4
    Firstwks = 7
    SecondWks = Firstwks + Tasksize + 1
    ThirdWks = SecondWks + Tasksize + 1
    FourthWks = ThirdWks + Tasksize + 1
    FifthWks = FourthWks + Tasksize + 1
    SixthWks = FifthWks + Tasksize + 1
End Enum

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Set ws = wb.Worksheets("Calendar")
    If Intersect(Target, ws.Range("B1")) Is Nothing Then Exit Sub
    
    Dim MonthSelected As Date, DoM As Date, CriteriaDate As Date
    Dim x As Integer, MonthNum As Integer
    Dim Opt_Events As String
    Dim wsE As Worksheet
    Dim lastRow As Long, y As Long
    Dim EventData As Variant
    Dim Dict_Holiday As Object
    Set Dict_Holiday = CreateObject("Scripting.Dictionary")
    TurnOffApp
    
    Set wsE = wb.Worksheets("Events")
    With wsE
        On Error Resume Next
        .ShowAllData
        On Error GoTo 0
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        EventData = .Range("A1:B" & lastRow).Value
        lastRow = .Cells(.Rows.Count, "L").End(xlUp).Row
        For y = 2 To lastRow Step 1
            If Not Dict_Holiday.Exists(.Cells(y, "L").Value) Then
                Dict_Holiday.Add .Cells(y, "L").Value, ""
            End If
        Next y
    End With
    

    With ws
        Opt_Events = .Cells(3, "I").Value
        MonthNum = Month(DateValue("01-" & .Range("B1") & "-2000"))
        MonthSelected = DateSerial(.Range("B2").Value, MonthNum, 1)
        DoM = MonthSelected - Weekday(MonthSelected, vbSunday)
        
        For x = 1 To 7 Step 1
            
            .Cells(RowWks.Firstwks, x).Value = DoM + x
            Call FilteredData(EventData, .Cells(RowWks.Firstwks, x), Opt_Events, MonthNum)
            .Cells(RowWks.SecondWks, x).Value = DoM + x + 7
            Call FilteredData(EventData, .Cells(RowWks.SecondWks, x), Opt_Events, MonthNum)
            .Cells(RowWks.ThirdWks, x).Value = DoM + x + 14
            Call FilteredData(EventData, .Cells(RowWks.ThirdWks, x), Opt_Events, MonthNum)
            .Cells(RowWks.FourthWks, x).Value = DoM + x + 21
            Call FilteredData(EventData, .Cells(RowWks.FourthWks, x), Opt_Events, MonthNum)
            .Cells(RowWks.FifthWks, x).Value = DoM + x + 28
            Call FilteredData(EventData, .Cells(RowWks.FifthWks, x), Opt_Events, MonthNum)
            .Cells(RowWks.SixthWks, x).Value = DoM + x + 35
            Call FilteredData(EventData, .Cells(RowWks.SixthWks, x), Opt_Events, MonthNum)
            If x > 1 And x < 7 Then
                Call FormatCells(.Cells(RowWks.Firstwks, x), MonthNum, Dict_Holiday)
                Call FormatCells(.Cells(RowWks.SecondWks, x), MonthNum, Dict_Holiday)
                Call FormatCells(.Cells(RowWks.ThirdWks, x), MonthNum, Dict_Holiday)
                Call FormatCells(.Cells(RowWks.FourthWks, x), MonthNum, Dict_Holiday)
                Call FormatCells(.Cells(RowWks.FifthWks, x), MonthNum, Dict_Holiday)
                Call FormatCells(.Cells(RowWks.SixthWks, x), MonthNum, Dict_Holiday)
            End If
        Next x
    End With
    TurnOnApp
End Sub
Sub FilteredData(EventData As Variant, rng As Range, Opt_Events As String, MonthNum As Integer)
    'CLEAR PREVIOUS DATA
    rng.Resize(RowWks.Tasksize, 1).Offset(1).ClearContents
    Dim CriteriaData As Date
    CriteriaData = rng.Value
    
    If Opt_Events = "Y" And Month(CriteriaData) <> MonthNum Then
        Exit Sub
    End If
    
    Dim coll As New Collection
    Dim i As Long
    For i = 1 To UBound(EventData) Step 1
        If EventData(i, 1) = CriteriaData Then
            coll.Add EventData(i, 2)
        End If
    Next i
    
    If coll.Count > 0 Then
        For i = 1 To coll.Count Step 1
            rng.Offset(i) = coll(i)
        Next i
    End If
End Sub

Sub FormatCells(rng As Range, MonthNum As Integer, Holidays As Object)
    Dim NewRng As Range
    Set NewRng = rng.Resize(RowWks.Tasksize + 1, 1)
    If Not Month(rng.Value) = MonthNum Then
        NewRng.Interior.Color = RGB(217, 217, 217) 'GREY FOR OTHER MONTHS
    ElseIf Holidays.Exists(rng.Value) Then
        NewRng.Interior.Color = RGB(217, 217, 217) 'GREY FOR HOLIDAYS
    Else
        NewRng.Interior.Color = RGB(255, 255, 255)
    End If
End Sub

