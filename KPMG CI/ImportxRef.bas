Option Explicit

Sub ImportxRef()
    Dim FilePath As String
    Dim wbName As String
    'OPEN FILE DIALOG TO SELECT THE FILE TO IMPORT
    With Application.FileDialog(msoFileDialogFilePicker)
        .InitialFileName = "*\*"
        .AllowMultiSelect = False
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb; *.csv", 1
        If .Show = -1 Then
            FilePath = .SelectedItems(1)
        Else
            MsgBox "xRef was not imported."
            Exit Sub
        End If
    End With
    
    Call TurnOffApp
    
    'CLEAR EXISTING DATA
    Dim cwb As Workbook
    Dim cws As Worksheet
    Dim lastRow As Integer
    
    Set cwb = ThisWorkbook
    Set cws = cwb.Worksheets("xRef")
    With cws
        lastRow = .Cells(.Rows.Count, 4).End(xlUp).Row
        .Range("A1:U" & lastRow + 3).Clear
    End With
    
    
    'COPY DATA FROM THE EXREF WORKSHEET
    Dim wb As Workbook
    Set wb = Workbooks.Open(FilePath, ReadOnly:=True)
    With wb.Worksheets("FA-Custody_xref")
        lastRow = .Cells(.Rows.Count, "E").End(xlUp).Row
        .Range("A1:T" & lastRow).Copy
    End With
    
    'PASTE THE COPIED DATA INTO THE WORKSHEET AND UPDATE THE OVERVIEW WORKSHEET WITH THE NAME OF THE IMPORT WORKBOOK
    With cws
        .Range("A1").PasteSpecial xlPasteAll
    End With
    ThisWorkbook.Worksheets("Macro").Cells(9, "B").Value = wb.Name
    'BREAK ALL LINKS
    On Error Resume Next
        Dim Arr_Links As Variant
        Dim y As Byte
        Arr_Links = cwb.LinkSources(Type:=xlLinkTypeExcelLinks)
        If Not IsEmpty(Arr_Links) Then
        For y = LBound(Arr_Links) To UBound(Arr_Links)
            cwb.BreakLink Name:=Arr_Links(y), Type:=xlLinkTypeExcelLinks
        Next y
        End If
    On Error GoTo 0
    
    Call TurnOnApp
    wb.Close False
    MsgBox "Completed"
End Sub
