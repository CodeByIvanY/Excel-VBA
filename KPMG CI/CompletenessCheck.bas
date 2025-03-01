Option Explicit
'Create By: Ivan Yang
'Department: DH CI
'Date: 3/1/2025

Sub CompletenessCheck()
    Dim cws As Worksheet, ws As Worksheet
    Dim wb As Workbook
    Dim FolderPath As String, FileName As String
    
    'PROMPT USER TO SELECT FOLDER CONTAINING CSV FILES
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder containing PBC files"
        .InitialFileName = ThisWorkbook.Path
        If .Show = -1 Then
            FolderPath = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With
    
    Call TurnOffApp

    'POPULATE FUND CODE DICTIONARY
    Dim dict_fundcode As Object
    Dim lastRow As Long, y As Long
    Dim FundCode As String
    Set cws = ThisWorkbook.Worksheets("xRef")
    Set dict_fundcode = CreateObject("Scripting.Dictionary")
    With cws
        lastRow = .Cells(.Rows.Count, "E").End(xlUp).Row
        For y = 6 To lastRow Step 1
            FundCode = .Cells(y, "D").Value
            If Not dict_fundcode.Exists(FundCode) Then
                dict_fundcode.Add FundCode, .Cells(y, "E").Value
            End If
        Next y
    End With
    
    'GET REPORT TYPE FROM MACRO SHEET
    Dim OptReport As String
    OptReport = ThisWorkbook.Worksheets("Macro").Cells(8, "B").Value

    
    FileName = Dir(FolderPath & "*.csv")
    Dim colNames As Variant
    Dim x As Byte
    Dim colNums As Variant
    Dim Rng As Range
    Dim RngFound As Range
    Dim clsEagle_data As clsEagle
    Dim clsNexen_data As clsNexen
    Dim dict_data As Object
    Dim str_colName As String
    Set dict_data = CreateObject("Scripting.Dictionary")
    
    'PROCESS EACH CSV FILE IN THE SELECTED FOLDER
    Do While FileName <> vbNullString
        Set wb = Workbooks.Open(FolderPath & FileName, ReadOnly:=True)
        colNames = GetColumnName(LCase(FileName))
        str_colName = colNames(0)
        If str_colName <> "N/A" Then
            Set ws = wb.Worksheets(1)
            ReDim colNums(1 To UBound(colNames)) As Variant
            Set Rng = ws.Rows(1)
            With Rng
                For x = 1 To UBound(colNames) Step 1
                    
                    Set RngFound = .Find(colNames(x), LookIn:=xlValues, Lookat:=xlWhole)
                    If Not RngFound Is Nothing Then
                        colNums(x) = RngFound.Column
                    Else
                        colNums(x) = -1
                    End If
                Next x
            End With
            lastRow = ws.Cells(ws.Rows.Count, colNums(1)).End(xlUp).Row
            
            If OptReport = "Nexen Reports" Then
            For y = 2 To lastRow
                FundCode = ws.Cells(y, colNums(1)).Value
                If dict_fundcode.Exists(FundCode) Then
                    FundCode = dict_fundcode(FundCode)
                End If

                ' Create or update clsNexen_data based on the column type
                If Not dict_data.Exists(FundCode) Then
                    Call CreateNexenData(dict_data, clsNexen_data, str_colName, ws, y, colNums, FileName, FundCode)
                    
                Else
                    Set clsNexen_data = dict_data(FundCode)
                    Call UpdateNexenData(clsNexen_data, str_colName, ws, y, colNums, FileName)
                End If
            Next y
            End If
            
            If OptReport = "Eagle Reports" Then
            For y = 2 To lastRow
                FundCode = ws.Cells(y, colNums(1)).Value
                If dict_fundcode.Exists(FundCode) Then
                    FundCode = dict_fundcode(FundCode)
                End If

                ' Create or update clsEagle_data based on the column type
                If Not dict_data.Exists(FundCode) Then
                    Call CreateEagleData(dict_data, clsEagle_data, str_colName, ws, y, colNums, FileName, FundCode)
                    
                Else
                    Set clsEagle_data = dict_data(FundCode)
                    Call UpdateEagleData(clsEagle_data, str_colName, ws, y, colNums, FileName)
                End If
            Next y
            End If
        End If
        wb.Close False
        FileName = Dir
    Loop
    
    'PREPARE EAGLE/NEXEN DATA FOR OUTPUT
    lastRow = dict_data.Count
    If OptReport = "Eagle Reports" Then
        x = 13
        ReDim Arr_data(1 To lastRow, 1 To x) As Variant
        For y = 1 To lastRow
            Set clsEagle_data = dict_data(dict_data.Keys()(y - 1))
            FillEagleDataArray Arr_data, y, clsEagle_data
        Next y
        Set cws = ThisWorkbook.Worksheets("Eagle")
        cws.Cells(2, 1).Resize(lastRow, x).Value = Arr_data
    ElseIf OptReport = "Nexen Reports" Then
        x = 18
        ReDim Arr_data(1 To lastRow, 1 To x) As Variant
        For y = 1 To lastRow
            Set clsNexen_data = dict_data(dict_data.Keys()(y - 1))
            FillNexenDataArray Arr_data, y, clsNexen_data
        Next y
        Set cws = ThisWorkbook.Worksheets("Nexen")
        cws.Cells(2, 1).Resize(lastRow, x).Value = Arr_data
    End If
    
    'CLEAN UP TIME FORMAT IN THE OUTPUT RANGE
    Set Rng = cws.Range(cws.Cells(2, 1), cws.Cells(lastRow, x))
    Rng.Replace "12:00:00 AM", Replacement:=""
    
    Call TurnOnApp
    
    MsgBox "Completed"
End Sub
Function GetColumnName(lcFileName As String) As Variant
    Select Case True
        Case InStr(1, lcFileName, "ledger", vbTextCompare) > 0 And InStr(1, lcFileName, "subledger", vbTextCompare) > 0             'Eagle Reports - Ledger Subledger
            GetColumnName = Array("Subledger", "Account", "Report End Date")
        Case InStr(1, lcFileName, "manual", vbTextCompare) > 0 And InStr(1, lcFileName, "ledger", vbTextCompare) > 0                'Eagle Reports - Manual Ledger
            GetColumnName = Array("Manual", "Account", "Report Start Date", "Report End Date")
        Case InStr(1, lcFileName, "unsettled", vbTextCompare) > 0 And InStr(1, lcFileName, "transactions", vbTextCompare) > 0       'Eagle Reports - Unsettled Transactions
            GetColumnName = Array("Unsettled", "Account/Sector Number", "Report End Date")
        Case InStr(1, lcFileName, "dividend", vbTextCompare) > 0 And InStr(1, lcFileName, "income", vbTextCompare) > 0              'Eagle Reports - Dividend Income
            GetColumnName = Array("Dividend", "Account/Sector", "Report Start Date", "Report End Date")
        Case InStr(1, lcFileName, "portfolio", vbTextCompare) > 0 And InStr(1, lcFileName, "valuation", vbTextCompare) > 0          'Eagle Reports - Portfolio Valuation
            GetColumnName = Array("Portfolio", "Sector ID", "Report Date")
        Case InStr(1, lcFileName, "asset", vbTextCompare) > 0 And InStr(1, lcFileName, "accrual", vbTextCompare) > 0                'Nexen Reports - Asset & Accrual
            GetColumnName = Array("Asset", "Master Number", "Effective Date")
        Case InStr(1, lcFileName, "class", vbTextCompare) > 0 And InStr(1, lcFileName, "trial", vbTextCompare) > 0                  'Nexen Reports - Class Trial Balance
            GetColumnName = Array("Class", "Master Number", "Begin Date", "End Date")
        Case InStr(1, lcFileName, "nexen", vbTextCompare) > 0 And InStr(1, lcFileName, "trial", vbTextCompare) > 0                  'Nexen Reports - Nexen Trial Balance
            GetColumnName = Array("Nexen", "Master Number", "Begin Date", "End Date")
        Case InStr(1, lcFileName, "distribution", vbTextCompare) > 0                                                                'Nexen Reports - Distribution Activity
            GetColumnName = Array("Distribution", "Account Number", "Begin Date", "End Date")
        Case InStr(1, lcFileName, "nav", vbTextCompare) > 0 And InStr(1, lcFileName, "summary", vbTextCompare) > 0                  'Nexen Reports - Nav Summary
            GetColumnName = Array("Nav", "Master Number", "Begin Date", "End Date")
        Case InStr(1, lcFileName, "trade", vbTextCompare) > 0 And InStr(1, lcFileName, "activity", vbTextCompare) > 0               'Nexen Reports - Trade Activity
            GetColumnName = Array("Trade", "Master Number", "Begin Date", "End Date")
        Case Else
            GetColumnName = Array("N/A")
    End Select
End Function

Sub FillEagleDataArray(Arr_data As Variant, index As Long, clsEagle_data As clsEagle)
    Arr_data(index, 1) = clsEagle_data.MasterCode
    Arr_data(index, 2) = clsEagle_data.Subledger_SourceFile
    Arr_data(index, 3) = clsEagle_data.Subledger_Date
    Arr_data(index, 4) = clsEagle_data.Manual_SourceFile
    Arr_data(index, 5) = clsEagle_data.Manual_BgnDate
    Arr_data(index, 6) = clsEagle_data.Manual_EndDate
    Arr_data(index, 7) = clsEagle_data.Unsettled_SourceFile
    Arr_data(index, 8) = clsEagle_data.Unsettled_Date
    Arr_data(index, 9) = clsEagle_data.Dividend_SourceFile
    Arr_data(index, 10) = clsEagle_data.Dividend_BgnDate
    Arr_data(index, 11) = clsEagle_data.Dividend_EndDate
    Arr_data(index, 12) = clsEagle_data.Portfolio_SourceFile
    Arr_data(index, 13) = clsEagle_data.Portfolio_Date
End Sub
Sub FillNexenDataArray(Arr_data As Variant, index As Long, clsNexen_data As clsNexen)
    Arr_data(index, 1) = clsNexen_data.MasterCode
    Arr_data(index, 2) = clsNexen_data.Asset_SourceFile
    Arr_data(index, 3) = clsNexen_data.Asset_Date
    Arr_data(index, 4) = clsNexen_data.ClassTB_SourceFile
    Arr_data(index, 5) = clsNexen_data.ClassTB_BgnDate
    Arr_data(index, 6) = clsNexen_data.ClassTB_EndDate
    Arr_data(index, 7) = clsNexen_data.NexenTB_SourceFile
    Arr_data(index, 8) = clsNexen_data.NexenTB_BgnDate
    Arr_data(index, 9) = clsNexen_data.NexenTB_EndDate
    Arr_data(index, 10) = clsNexen_data.Distribution_SourceFile
    Arr_data(index, 11) = clsNexen_data.Distribution_BgnDate
    Arr_data(index, 12) = clsNexen_data.Distribution_EndDate
    Arr_data(index, 13) = clsNexen_data.Nav_SourceFile
    Arr_data(index, 14) = clsNexen_data.Nav_BgnDate
    Arr_data(index, 15) = clsNexen_data.Nav_EndDate
    Arr_data(index, 16) = clsNexen_data.Trade_SourceFile
    Arr_data(index, 17) = clsNexen_data.Trade_BgnDate
    Arr_data(index, 18) = clsNexen_data.Trade_EndDate
End Sub
Sub UpdateEagleData(clsEagle_data As clsEagle, colName As String, ws As Worksheet, rowIndex As Long, colNums As Variant, FileName As String)
    Dim date_1 As Date, date_2 As Date
    Select Case colName
        Case "Subledger"
            clsEagle_data.Subledger_SourceFile = FileName
            date_1 = ws.Cells(rowIndex, colNums(2)).Value
            If clsEagle_data.Subledger_Date < date_1 Then
                clsEagle_data.Subledger_Date = date_1
            End If
        Case "Manual"
            clsEagle_data.Manual_SourceFile = FileName
            date_1 = ws.Cells(rowIndex, colNums(2)).Value
            date_2 = ws.Cells(rowIndex, colNums(3)).Value
            If clsEagle_data.Manual_BgnDate > date_1 Or clsEagle_data.Manual_BgnDate <= 0 Then
                clsEagle_data.Manual_BgnDate = date_1
            End If
            If clsEagle_data.Manual_EndDate < date_2 Then
                clsEagle_data.Manual_EndDate = date_2
            End If
        Case "Unsettled"
            clsEagle_data.Unsettled_SourceFile = FileName
            date_1 = ws.Cells(rowIndex, colNums(2)).Value
            If clsEagle_data.Unsettled_Date < date_1 Then
                clsEagle_data.Unsettled_Date = date_1
            End If
        Case "Dividend"
            clsEagle_data.Dividend_SourceFile = FileName
            date_1 = ws.Cells(rowIndex, colNums(2)).Value
            date_2 = ws.Cells(rowIndex, colNums(3)).Value
            If clsEagle_data.Dividend_BgnDate > date_1 Or clsEagle_data.Dividend_BgnDate <= 0 Then
                clsEagle_data.Dividend_BgnDate = date_1
            End If
            If clsEagle_data.Dividend_EndDate < date_2 Then
                clsEagle_data.Dividend_EndDate = date_2
            End If
        Case "Portfolio"
            clsEagle_data.Portfolio_SourceFile = FileName
            date_1 = ws.Cells(rowIndex, colNums(2)).Value
            If clsEagle_data.Portfolio_Date < date_1 Then
                clsEagle_data.Portfolio_Date = date_1
            End If
    End Select
End Sub
Sub UpdateNexenData(clsNexen_data As clsNexen, colName As String, ws As Worksheet, rowIndex As Long, colNums As Variant, FileName As String)
    Dim date_1 As Date, date_2 As Date
    Select Case colName
        Case "Asset"
            clsNexen_data.Asset_SourceFile = FileName
            date_1 = ws.Cells(rowIndex, colNums(2)).Value
            If clsNexen_data.Asset_Date < date_1 Then
                clsNexen_data.Asset_Date = date_1
            End If
        Case "Class"
            clsNexen_data.ClassTB_SourceFile = FileName
            date_1 = ws.Cells(rowIndex, colNums(2)).Value
            date_2 = ws.Cells(rowIndex, colNums(3)).Value
            If clsNexen_data.ClassTB_BgnDate > date_1 Or clsNexen_data.ClassTB_BgnDate <= 0 Then
                clsNexen_data.ClassTB_BgnDate = date_1
            End If
            If clsNexen_data.ClassTB_EndDate < date_2 Then
                clsNexen_data.ClassTB_EndDate = date_2
            End If
        Case "Nexen"
            clsNexen_data.NexenTB_SourceFile = FileName
            date_1 = ws.Cells(rowIndex, colNums(2)).Value
            date_2 = ws.Cells(rowIndex, colNums(3)).Value
            If clsNexen_data.NexenTB_BgnDate > date_1 Or clsNexen_data.NexenTB_BgnDate <= 0 Then
                clsNexen_data.NexenTB_BgnDate = date_1
            End If
            If clsNexen_data.NexenTB_EndDate < date_2 Then
                clsNexen_data.NexenTB_EndDate = date_2
            End If
        Case "Distribution"
            clsNexen_data.Distribution_SourceFile = FileName
            date_1 = ws.Cells(rowIndex, colNums(2)).Value
            date_2 = ws.Cells(rowIndex, colNums(3)).Value
            If clsNexen_data.Distribution_BgnDate > date_1 Or clsNexen_data.Distribution_BgnDate <= 0 Then
                clsNexen_data.Distribution_BgnDate = date_1
            End If
            If clsNexen_data.Distribution_EndDate < date_2 Then
                clsNexen_data.Distribution_EndDate = date_2
            End If
        Case "Nav"
            clsNexen_data.Nav_SourceFile = FileName
            date_1 = ws.Cells(rowIndex, colNums(2)).Value
            date_2 = ws.Cells(rowIndex, colNums(3)).Value
            If clsNexen_data.Nav_BgnDate > date_1 Or clsNexen_data.Nav_BgnDate <= 0 Then
                clsNexen_data.Nav_BgnDate = date_1
            End If
            If clsNexen_data.Nav_EndDate < date_2 Then
                clsNexen_data.Nav_EndDate = date_2
            End If
        Case "Trade"
            clsNexen_data.Trade_SourceFile = FileName
            date_1 = ws.Cells(rowIndex, colNums(2)).Value
            date_2 = ws.Cells(rowIndex, colNums(3)).Value
            If clsNexen_data.Trade_BgnDate > date_1 Or clsNexen_data.Trade_BgnDate <= 0 Then
                clsNexen_data.Trade_BgnDate = date_1
            End If
            If clsNexen_data.Trade_EndDate < date_2 Then
                clsNexen_data.Trade_EndDate = date_2
            End If
    End Select
End Sub
Sub CreateEagleData(dict_data As Object, clsEagle_data As clsEagle, colName As String, ws As Worksheet, rowIndex As Long, colNums As Variant, FileName As String, FundCode As String)
    Set clsEagle_data = New clsEagle
    clsEagle_data.MasterCode = FundCode
    Select Case colName
        Case "Subledger"
            clsEagle_data.Subledger_SourceFile = FileName
            clsEagle_data.Subledger_Date = ws.Cells(rowIndex, colNums(2)).Value
        Case "Manual"
            clsEagle_data.Manual_SourceFile = FileName
            clsEagle_data.Manual_BgnDate = ws.Cells(rowIndex, colNums(2)).Value
            clsEagle_data.Manual_EndDate = ws.Cells(rowIndex, colNums(3)).Value
        Case "Unsettled"
            clsEagle_data.Unsettled_SourceFile = FileName
            clsEagle_data.Unsettled_Date = ws.Cells(rowIndex, colNums(2)).Value
        Case "Dividend"
            clsEagle_data.Dividend_SourceFile = FileName
            clsEagle_data.Dividend_BgnDate = ws.Cells(rowIndex, colNums(2)).Value
            clsEagle_data.Dividend_EndDate = ws.Cells(rowIndex, colNums(3)).Value
        Case "Portfolio"
            clsEagle_data.Portfolio_SourceFile = FileName
            clsEagle_data.Portfolio_Date = ws.Cells(rowIndex, colNums(2)).Value
    End Select
    dict_data.Add FundCode, clsEagle_data
End Sub
Sub CreateNexenData(dict_data As Object, clsNexen_data As clsNexen, colName As String, ws As Worksheet, rowIndex As Long, colNums As Variant, FileName As String, FundCode As String)
    Set clsNexen_data = New clsNexen
    clsNexen_data.MasterCode = FundCode
    Select Case colName
        Case "Asset"
            clsNexen_data.Asset_SourceFile = FileName
            clsNexen_data.Asset_Date = ws.Cells(rowIndex, colNums(2)).Value
        Case "Class"
            clsNexen_data.ClassTB_SourceFile = FileName
            clsNexen_data.ClassTB_BgnDate = ws.Cells(rowIndex, colNums(2)).Value
            clsNexen_data.ClassTB_EndDate = ws.Cells(rowIndex, colNums(3)).Value
        Case "Nexen"
            clsNexen_data.NexenTB_SourceFile = FileName
            clsNexen_data.NexenTB_BgnDate = ws.Cells(rowIndex, colNums(2)).Value
            clsNexen_data.NexenTB_EndDate = ws.Cells(rowIndex, colNums(3)).Value
        Case "Distribution"
            clsNexen_data.Distribution_SourceFile = FileName
            clsNexen_data.Distribution_BgnDate = ws.Cells(rowIndex, colNums(2)).Value
            clsNexen_data.Distribution_EndDate = ws.Cells(rowIndex, colNums(3)).Value
        Case "Nav"
            clsNexen_data.Nav_SourceFile = FileName
            clsNexen_data.Nav_BgnDate = ws.Cells(rowIndex, colNums(2)).Value
            clsNexen_data.Nav_EndDate = ws.Cells(rowIndex, colNums(3)).Value
        Case "Trade"
            clsNexen_data.Trade_SourceFile = FileName
            clsNexen_data.Trade_BgnDate = ws.Cells(rowIndex, colNums(2)).Value
            clsNexen_data.Trade_EndDate = ws.Cells(rowIndex, colNums(3)).Value
    End Select
    dict_data.Add FundCode, clsNexen_data
End Sub

