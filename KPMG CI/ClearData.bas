Option Explicit

Sub ClearData()
    Dim ws As Worksheet, wsMacro As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Dim OptClear As String
    Dim lastRow As Integer
    
    Call TurnOffApp
    
    Set wsMacro = wb.Worksheets("Macro")
    OptClear = wsMacro.Cells(7, "B").Value
    
    If OptClear = "Eagle worksheet" Then
    Else
        Set ws = wb.Worksheets("Nexen")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 3
        ws.Rows("2:" & lastRow).Clear
    End If
    If OptClear = "Nexen worksheet" Then
    Else
        Set ws = wb.Worksheets("Eagle")
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 3
        ws.Rows("2:" & lastRow).Clear
    End If
    
    Call TurnOnApp
    MsgBox "Data cleared successfully from " & OptClear & "!"
    
End Sub
