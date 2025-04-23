Option Explicit
'CREATED BY: IVAN Y - CI
'DATE: 04/22/2025
'PURPOSE: TO READ EACH FILE IN THE ZIP AND LIST IN THE WORKSHEET
Sub ListZipDetails()
    Dim y As Long
    Dim PathFilename As Variant, FileNameInZip As Variant
    
    'PROMPT THE USER TO SELECT A ZIP FILE
    PathFilename = Application.GetOpenFilename("ZipFiles (*.zip), *.zip")
    If PathFilename = "False" Then Exit Sub
    
    'CLEAR PREIVOUS DATA IN COLUMN A
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Zipped Check")
    With ws
        y = .Cells(.Rows.Count, "A").End(xlUp).Row
        .Range("A2:A" & y + 5).Clear
        y = 1
    End With
  
    'LOOP THROUGH EACH FILE IN THE ZIP AND LIST IN THE WORKSHEET
    Dim oApp As Object
    Set oApp = CreateObject("Shell.Application")
    On Error Resume Next
    For Each FileNameInZip In oApp.Namespace(PathFilename).Items
        If Not FileNameInZip Is Nothing Then
            y = y + 1
            ws.Cells(y, "A").Value = FileNameInZip
        End If
    Next
    
    Set oApp = Nothing
    MsgBox "Completed", vbInformation
End Sub
