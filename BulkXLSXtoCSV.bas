Option Explicit
'Author: IVAN Y
'Date Created: 12/31/2024
'Date Modified: N/A
'Purpose: To convert XSLX files to CSV
Sub BulkXLSXtoCSV()
    Dim xlsxFolder As String
    Dim fileName As String, lcFileName As String
    Dim wb As Workbook
    Dim wsName As Variant
    Dim ws As Worksheet
    Dim x As Integer
       
    'SELECT THE FOLDER WHICH CONTAINS THE XLSX FILES
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select folder containg XLSX files"
        .InitialFileName = ActiveWorkbook.Path
        If .Show Then
            xlsxFolder = .SelectedItems(1) & "\"
        Else
            Exit Sub
        End If
    End With
    
    'TURN OFF SCREEN UPDATING AND ALTERS FOR SMOOTHER OPERATION
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    'DETERMINE THE FILES IN THE FOLDER
    fileName = Dir(xlsxFolder & "*.xlsx")
    'LOOP THROUGH THE FILES IN THE FOLDER
    Do While fileName <> vbNullString
        Set wb = Workbooks.Open(xlsxFolder & fileName, ReadOnly:=True)
        'DETERMINE THE WORKSHEET TO COPY BASED ON THE FILE NAME:
        wsName = GetWorksheetName(LCase(fileName))
        
        'COPY THE DETERMINED WORKSHEET AND SAVE AS CSV
        On Error Resume Next ' HANDLE POTENTIAL ERRORS IF THE WORKSHEET DOES NOT EXIST 
            For x = LBound(wsName) To UBound(wsName) Step 1
                If wsName(x) = "Default" Then
                    Set ws = wb.Worksheets(1)
                Else
                    Set ws = wb.Worksheets(wsName(x))
                End If
                If Not ws Is Nothing Then
                    ws.Copy
                    ActiveWorkbook.SaveAs fileName:=xlsxFolder & Replace(fileName, ".xlsx", " - " & ws.Name & ".csv", vbTextCompare), FileFormat:=xlCSV
                    ActiveWorkbook.Close False
                Else
                    MsgBox "Worksheet '" & wsName & "' not found in " & fileName, vbExclamation
                End If
            Next x
        On Error GoTo 0 ' RESET ERROR HANDLING
        
        wb.Close False
        fileName = Dir
    Loop
    
    'TURN ALTERS AND SCREEN UPDATING BACK ON
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "The files have been converted. Thank you!"

End Sub
Function GetWorksheetName(lcFileName As String) As Variant
    If InStr(1, lcFileName, "asset", vbTextCompare) > 0 And InStr(1, lcFileName, "accrual", vbTextCompare) > 0 Then
        GetWorksheetName = Array("Asset", "Asset23")
    ElseIf InStr(1, lcFileName, "nav", vbTextCompare) > 0 And InStr(1, lcFileName, "summary", vbTextCompare) > 0 Then
        GetWorksheetName = Array("Nav")
    ElseIf InStr(1, lcFileName, "unsettled", vbTextCompare) > 0 And InStr(1, lcFileName, "transactions", vbTextCompare) > 0 Then
        GetWorksheetName = Array("unsettled")
    Else
        GetWorksheetName = Array("Default") 'DEFAULT TO THE FIRST WORKSHEET IF NO CONDITIONS ARE MET
    End If
End Function


