Option Explicit
'CREATED BY: IVAN Y - CI
'UPDATED DATE: 2025-03-24
'PURPOSE: To AUTOMATE THE PROCESS OF MOVING OR COPYING EXCEL FILES FROM A SPECIFIED SOURCE FOLDER TO A DESTINATION FOLDER BASED ON A LIST OF FUND CODES PROVIDED IN AN EXCEL WORKSHEET.

Sub MoveXlsxFiles()
    Dim fso As Object
    Dim sourceFolder As String
    Dim destinationFolder As String
    Dim file As Object
    Dim sourcePath As String
    Dim destPath As String
    Dim dict_fund As Object
    Dim ws As Worksheet
    Dim lastRow As Integer, y As Integer
    Dim FundCode As String
    
    Set ws = ThisWorkbook.Worksheets("Move Copy Funds")
    Set dict_fund = CreateObject("Scripting.Dictionary")
    
    'SET THE SOURCE AND DESTINATION FOLDER PATHS
    Dim File_Extension As String
    Dim FundCode_Type As String
    Dim MoveType As String
    Dim ActionType As String
    
    
    With ws
        'SET THE SOURCE AND DESTINATION FOLDER PATHS
        File_Extension = LCase(.Cells(4, "B").Value)
        FundCode_Type = Left(.Cells(5, "B").Value, 6)
        ActionType = .Cells(3, "B").Value
        If ActionType <> "MOVE" And ActionType <> "COPY" Then
            MsgBox "Invalid Action specified. Please use 'Move' or 'Copy'.", vbExclamation
            Exit Sub
        End If
        
        sourceFolder = .Cells(1, 2).Value & "\"
        destinationFolder = .Cells(2, 2).Value & "\"
        
        'POPULATE THE DICTIONARY WITH FUND CODES
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        For y = 7 To lastRow Step 1
            FundCode = .Cells(y, 1).Value
            If Not dict_fund.Exists(FundCode) Then
                dict_fund.Add FundCode, ""
            Else
                MsgBox FundCode & " already exists in the list; Please verify the list.", vbExclamation
                Exit Sub
            End If
        Next y
    End With
    
    'CREATE A FILESYSTEMOBJECT
    Set fso = CreateObject("Scripting.FileSystemObject")
    'CHECK IF THE SOURCE FOLDER EXISTS
    If Not fso.FolderExists(sourceFolder) Then
        MsgBox "SOURCE FOLDER DOES NOT EXIST: " & sourceFolder
        Exit Sub
    End If
    'LOOP THROUGH EACH FILE IN THE SOURCE FOLDER
    For Each file In fso.GetFolder(sourceFolder).Files
        'DETERMINE THE FUND CODE BASED ON THE FILE NAME
        FundCode = GetFundCode(file.Name, FundCode_Type)
        
        If LCase(fso.GetExtensionName(file.Name)) = File_Extension And dict_fund.Exists(FundCode) Then
            sourcePath = file.Path
            destPath = destinationFolder & file.Name
            'MOVE OR COPY THE FILE TO THE DESTINATION FOLDER
            If ActionType = "MOVE" Then
                fso.MoveFile sourcePath, destPath
            ElseIf ActionType = "COPY" Then
                fso.CopyFile sourcePath, destPath
            End If
        End If
    Next file
    MsgBox "Files have been moved or copied!"
End Sub
'FUNCTION TO DETERMINE THE FUND CODE BASED ON THE FILE NAME AND FUND CODE TYPE
Public Function GetFundCode(fileName As String, fundCodeType As String) As String
    If fundCodeType = "Type 2" Then
        GetFundCode = Trim(Split(fileName, " - ")(1))
    Else
        If Left(fileName, 2) = "UF" Or Left(fileName, 3) = "UIF" Then
            GetFundCode = "UF"
        ElseIf InStr(1, UCase(Left(fileName, 11)), "_CAD", vbTextCompare) > 0 Then
            GetFundCode = Left(fileName, 11)
        Else
            GetFundCode = Left(fileName, 7)
        End If
    End If
End Function
