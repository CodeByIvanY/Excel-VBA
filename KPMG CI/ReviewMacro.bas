Option Explicit
'UPDATED BY: IVAN Y - CI
Sub FileCheck()
    'Declare variable
    Dim ws As Worksheet
    Dim y As Integer, openPos As Integer, closePos As Integer, lastRow As Integer, lastRowX As Integer
    Dim FundCode As String, fileName As String
    Dim folderpath As String
    Dim FileExtension As String
    Dim FundCode_Type As String
    'Set ws to the "Review Macro"worksheet
    Set ws = ThisWorkbook.Worksheets("Review Macro")
    
    'Get the folder path from cell E6 and append a backslash
    With ws
        folderpath = .Cells(6, 5).Value & "\"
        FileExtension = .Cells(4, "E").Value
        FundCode_Type = Left(.Cells(5, "E").Value, 6)
        lastRow = .Cells(Rows.Count, 2).End(xlUp).Row
        .Range("A5:B" & lastRow + 5).Clear
    End With
    
    
    'Initialize y to 4
    y = 4
    
    'Get the name of the first read-only Excel file in the folder
    fileName = Dir(folderpath & "*" & FileExtension, vbReadOnly = True)
    
    'Loop as long as there are more files to process
    Do While fileName <> ""
    ' Increment y by 1
        y = y + 1
        ' Write the filename to cell B(y) & Fundcode to cell A(y)
        With ws
            .Cells(y, 1) = GetFundCode(fileName, FundCode_Type)
            .Cells(y, 2) = fileName
        End With
        
        ' Get the name of the next file
        fileName = Dir
    Loop
    'Write the number of files processed to cell E7
    With ws
        .Cells(7, 5).Value = y - 4
    End With
    
End Sub

'This VBA code is designed to automate the process of copying specific worksheets and ranges from workbook [review_template] to individual fund export.
Sub vba_copy_sheet()
    Dim wb_Review As Workbook, wb_Fund As Workbook
    Dim cws As Worksheet, ws As Worksheet, RTws As Worksheet, ws_FOFControlled As Worksheet
    Dim fileName As String, FundCode As String, FundType As String, sourceFilename As String
    Dim APath As String, BPath As String, CPath As String
    Dim lastRow As Long, lastRowFOF As Long, lastRowTC As Long, i As Long
    Dim FindX As String, TCRange As String
    Dim Arr_ws As Variant
    Dim StartTime As Double
    Dim lastRowType As Integer

    StartTime = Timer
    
    'TURN OFF EXCEL SETTINGS FOR FASTER EXECUTION
    Call TurnOffApp
    'CLEAR SPECIFIC RANGE IN THE "REVIEW MACRO" AND SET THE PATHS FOR THE SOURCE AND DESTINATION FILES
    Set cws = ThisWorkbook.Worksheets("Review Macro")
    With cws
        lastRow = .Cells(Rows.Count, 2).End(xlUp).Row
        .Range("A5:B" & lastRow + 5).Clear
        APath = .Range("E16").Value & "\"
        BPath = .Range("E6").Value & "\"
        CPath = .Range("E17").Value & "\"
        FundType = .Range("E14").Value
        sourceFilename = .Range("E15").Value
        FindX = "[" & sourceFilename & ".xlsx" & "]"
    End With

    'CREATE AN ARRAY FOR THE NAMES OF THE WORKSHEETS TO COPY
    If FundType <> "MFC" Then
        Arr_ws = Array("Last Distribution Tax Calc", "Review", "Derivatives", "Adjustment Summary", "TaxInputsheet", "Tax Calculation", "Allocation Updated", "Allocation Income - Class", "Allocation - Gain - Class", "CG Inclusion Details")
        TCRange = "O1:U500"
        lastRowType = 274
    Else
        Arr_ws = Array("Last Distribution Tax Calc", "Review", "Derivatives", "Adjustment Summary", "TaxInputsheet", "Tax Calculation", "CG Inclusion Details")
        TCRange = "N1:T233"
        lastRowType = 236
    End If

    'OPEN THE SOURCE WORKBOOK
    Set wb_Review = Workbooks.Open(APath & sourceFilename & ".xlsx", ReadOnly:=True)
    'SET REFERENCES TO SPECIFIC WORKSHEETS AND GET LAST ROWS
    Set RTws = wb_Review.Worksheets("LDTC")
    lastRowTC = RTws.Cells(RTws.Rows.Count, 1).End(xlUp).Row
    Set ws_FOFControlled = wb_Review.Worksheets("FOF_Controlled")
    With ws_FOFControlled
        On Error Resume Next
        .ShowAllData
        On Error GoTo 0
        lastRowFOF = ws_FOFControlled.Cells(ws_FOFControlled.Rows.Count, 1).End(xlUp).Row
    End With
    'LOOP THROUGH ALL .XLSX FILES IN TEH DESIGNATION FOLDER
    fileName = Dir(BPath & "*.xlsx")
    Do While fileName <> ""
        Set wb_Fund = Workbooks.Open(BPath & fileName)
        FundCode = Trim(Split(fileName, "-")(1))

        'COPY SPECIFIED WORKSHEETS FROM THE SOURCE WORKBOOK TO THE DESIGNATION WORKBOOK
        With wb_Review
            .Sheets(Array("Last Distribution Tax Calc", "Review", "Derivatives", "Adjustment Summary", "TaxInputsheet")).Copy After:=wb_Fund.Sheets("Tax Calculation")
        End With
        
        ' COPY SPECIFIC RANGES FROM SOURCE TO DESINTATION
        CopyRange wb_Review.Sheets("Tax Calculation"), wb_Fund.Sheets("Tax Calculation"), TCRange
        
        If FundType <> "MFC" Then
            CopyRange wb_Review.Sheets("Allocation Updated"), wb_Fund.Sheets("Allocation Updated"), "T1:AZ500"
            CopyRange wb_Review.Sheets("Allocation Income - Class"), wb_Fund.Sheets("Allocation Income - Class"), "A140:AZ500"
            CopyRange wb_Review.Sheets("Allocation - Gain - Class"), wb_Fund.Sheets("Allocation - Gain - Class"), "A140:AZ500"
        End If
        ' UPDATE DATE FORMATS IN SPECIFIC SHEETS FOR CG INCLUSION DETAILS
        'UpdateDateFormats wb_Fund.Sheets("Detailed Tax Gain Loss"), "J", "I"
        'UpdateDateFormats wb_Fund.Sheets("Stop Loss Detailed"), "H"
        'UpdateDateFormats wb_Fund.Sheets("Suspended Loss PR Detailed"), "E"
        'CopyRange wb_Review.Sheets("CG Inclusion Details"), wb_Fund.Sheets("CG Inclusion Details"), "D1:H100"

        
        'LAST DISTRIBUTION TAX CALC DATA UPDATE
        Call LastDistTaxCal(wb_Fund.Worksheets("Last Distribution Tax Calc"), RTws, FundCode, lastRowTC, lastRowType)
        
        'UPDATE THE "FOF" WORHSEET IN THE DESTINATION WORKBOOK WITH DATA FROM REVIEW TEMPLATE BASED ON A SPECIFIC FUND CODE.
        On Error Resume Next
        Call FOF_Controlled_Check(wb_Fund.Sheets("FOF Controlled Summary"), ws_FOFControlled, lastRowFOF, FundCode)
        On Error GoTo 0
        '****FORMULA UPDATING****
        'REPLACE REFERENCES TO THE SOURCE WORKBOOK IN THE COPIED WORKSHEET.
        For i = LBound(Arr_ws) To UBound(Arr_ws) Step 1
            Set ws = wb_Fund.Worksheets(Arr_ws(i))
            With ws
                .Cells.Replace What:=FindX, Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False
            End With
        Next i
        
        'SAVE AND CLOSE THE DESTINATION WORKBOOK.
        wb_Fund.SaveAs fileName:=CPath & fileName & "-review.xlsx", FileFormat:=51
        wb_Fund.Close
        
        ' WRITE FILENAME AND FUND CODE TO "REVIEW MACRO" WORKSHEET.
        With cws
            lastRow = .Cells(Rows.Count, 2).End(xlUp).Row + 1
            .Cells(lastRow, 2).Value = fileName
            .Cells(lastRow, 1).Value = FundCode
        End With
        
        'GET NEXT FILE
        fileName = Dir
    Loop
    'CLOSE REVIEW TEMPLATE WORKBOOK
    wb_Review.Close
    
    
    ' TURN EXCEL SETTINGS BACK ON.
    Call TurnOnApp
    If cws.Cells(18, "E").Value = "Yes" Then
        Call EmailSelf
    End If
    Application.Wait (Now + TimeValue("0:00:05"))
    ' Display message box indicating successful completion and runtime
    MsgBox "Copied sheets & formulas successfully -  RunTime : " & Format((Timer - StartTime) / 86400, "hh:mm:ss")

End Sub
'HELPER FUNCTION - TO FILTER AND PROCESS DATA BASED ON A SPECIFIED FUND CODE, PERFORMING CROSS-CHECK AND APPLYING CONDITIONAL FORMATTING.
Sub FOF_Controlled_Check(ws As Worksheet, ws_FOFControlled As Worksheet, lastRowFOF As Long, FundCode As String)
    Dim Int_X As Integer, Int_Y As Integer, Int_Z As Integer, Int_Temp As Integer, lastRow As Integer
    Dim Condition1 As FormatCondition, Condition2 As FormatCondition, Condition3 As FormatCondition, Condition4 As FormatCondition
    Dim rng As Range
    
    On Error Resume Next
    'FILTER THE FOFCONTROLLED WORKSHEET BASED ON THE FUNDCODE
    With ws_FOFControlled
        .ShowAllData
        .Range("A1:AI" & lastRowFOF).AutoFilter Field:=1, Criteria1:=FundCode
        .Range("B1:AI" & lastRowFOF).SpecialCells(xlCellTypeVisible).Copy
    End With
    
    'PASTE THE FILTERED DATA INTO THE TARGET WORKSHEET
    With ws
        Int_X = .Cells(.Rows.Count, 1).End(xlUp).Row + 2
        .Cells(Int_X, "A").PasteSpecial Paste:=xlPasteAll
        Int_Y = .Cells(.Rows.Count, 1).End(xlUp).Row + 2
        .Cells(Int_Y, "A").PasteSpecial Paste:=xlPasteAll
        
        'APPLY FORMULAS TO THE NEW ROWS
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        Int_Temp = Int_Y - Int_X
        For Int_Z = Int_Y + 1 To lastRow Step 1
            .Cells(Int_Z, "R").Formula2 = "=IFERROR(XMATCH($A" & Int_Z & ",$A$1:$A$" & Int_X & "),""FALSE"")"
            .Cells(Int_Z, "M").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,12)-M" & Int_Z - Int_Temp
                
            .Cells(Int_Z, "S").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,18)-S" & Int_Z - Int_Temp
            .Cells(Int_Z, "T").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,19)-T" & Int_Z - Int_Temp
            .Cells(Int_Z, "U").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,20)-U" & Int_Z - Int_Temp
            .Cells(Int_Z, "V").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,21)-V" & Int_Z - Int_Temp
            .Cells(Int_Z, "W").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,22)-W" & Int_Z - Int_Temp
            .Cells(Int_Z, "X").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,23)-X" & Int_Z - Int_Temp
                
            .Cells(Int_Z, "Z").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,25)-Z" & Int_Z - Int_Temp
            .Cells(Int_Z, "AA").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,26)-AA" & Int_Z - Int_Temp
            .Cells(Int_Z, "AB").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,27)-AB" & Int_Z - Int_Temp
            .Cells(Int_Z, "AC").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,28)-AC" & Int_Z - Int_Temp
            .Cells(Int_Z, "AD").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,29)-AD" & Int_Z - Int_Temp
            .Cells(Int_Z, "AE").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,30)-AE" & Int_Z - Int_Temp
            .Cells(Int_Z, "AF").Formula2 = "=OFFSET($A$1,R" & Int_Z & "-1,31)-AF" & Int_Z - Int_Temp
        Next Int_Z
        
        .Cells(10, "R").Value = "Cross Check"
        .Cells(10, "R").Interior.Color = RGB(255, 192, 0)
        
        Int_X = Int_X - 3
        For Int_Z = 11 To Int_X Step 1
            .Cells(Int_Z, "R").Formula2 = "=ISNUMBER(XMATCH($A" & Int_Z & ",$A$" & Int_X + 4 & ":$A$" & Int_Y - 2 & "))"
        Next Int_Z
        
        'APPLY CONDITIONAL FORMATTING
        Set Condition1 = .Range("R11:R" & lastRow).FormatConditions.Add(Type:=xlTextString, String:="FALSE", TextOperator:=xlContains)
        Set rng = .Range("M" & Int_Y + 1 & ":M" & lastRow & ",S" & Int_Y + 1 & ":AF" & lastRow)
        Set Condition2 = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="1")
        Set Condition3 = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="-1")
        Set Condition4 = rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISERROR(M" & Int_Y + 1 & ")")
    End With
    'SET FORMATTING COLOURS TO RED
    With Condition1
        .Interior.Color = vbRed
    End With
    With Condition2
        .Interior.Color = vbRed
    End With
    With Condition3
        .Interior.Color = vbRed
    End With
    With Condition4
        .Interior.Color = vbRed
    End With
    On Error GoTo 0
End Sub
'HELPER FUNCTION TO COPY LASTDISTRIBUTIONTAXCALC DATA TO RESPECTIVE WORKSHEET BASED ON A SPECIFIED FUND CODE
Sub LastDistTaxCal(ws_LDTC As Worksheet, RTws As Worksheet, FundCode As String, lastRowTC As Long, lastRowType As Integer)
    Dim Arr_desc() As Variant, Arr_data() As Variant
    ReDim Arr_desc(1 To lastRowType, 1 To 1) As Variant, Arr_data(1 To lastRowType, 1 To 12) As Variant
    On Error Resume Next
    With RTws
        .ShowAllData
        .Range("A1:O" & lastRowTC).AutoFilter Field:=2, Criteria1:=FundCode
        Arr_desc = .Range("C2:C" & lastRowTC).SpecialCells(xlCellTypeVisible)
        Arr_data = .Range("D2:O" & lastRowTC).SpecialCells(xlCellTypeVisible)
    End With
        
    With ws_LDTC
        .Range("A1").Resize(lastRowType, 1).Value = Arr_desc
        .Range("C1").Resize(lastRowType, 12).Value = Arr_data
        .Range("B7").Value = FundCode
    End With
    On Error GoTo 0
End Sub
'HELPER FUNCTION TO COPY A RANGE FROM ONE WORKSHEET TO ANOTHER
Sub CopyRange(sourceSheet As Worksheet, destSheet As Worksheet, rng As String)
    sourceSheet.Range(rng).Copy
    destSheet.Range(rng).PasteSpecial xlPasteAll
End Sub

'HELPER FUNCTION TO UPDATE DATE FORMATS IN SPECIFIED COLUMNS
Sub UpdateDateFormats(ws As Worksheet, ParamArray columns() As Variant)
    Dim lastRow As Long, i As Long, j As Long
    For i = LBound(columns) To UBound(columns)
        Dim col As String
        col = columns(i)
        lastRow = ws.Cells(Rows.Count, col).End(xlUp).Row
        For j = 3 To lastRow
            If IsNumeric(Right(ws.Range(col & j).Value, 4)) Then
                ws.Range(col & j).Value = CDate(ws.Range(col & j).Value)
            End If
        Next j
    Next i
End Sub
Sub EmailSelf()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim EmailAddress As String
    'TURN OFF APPLICATION UPDATES OR ALERTS
    Call TurnOffApp
    'CREATE OUTLOOK APPLICATION AND MAIL ITEM
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    'GET THE EMAIL ADDRESS FROM THE SPECIFIED CELL
    EmailAddress = ThisWorkbook.Worksheets("Review Macro").Cells(19, "E").Value
    On Error Resume Next
        With OutlookMail
        .To = EmailAddress
        .Subject = "CI Review Template Macro Completed"
        .Body = "Completed"
        .Display
        End With
        Application.Wait (Now + TimeValue("0:00:02"))
        Application.SendKeys "%{s}", True
    On Error GoTo 0
    Set OutlookApp = Nothing
    Set OutlookMail = Nothing
    Call TurnOnApp
End Sub

