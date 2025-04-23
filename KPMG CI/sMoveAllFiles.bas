Option Explicit
'CREATED BY: IVAN Y - CI
'PURPOSE: TO TRANSFER ALL FILES LOCATED WITHIN THE SUBFOLDERS OF BULK EXPORT\CI. THE CODE SYSTEMATICALLY NAVIGATES THROUGH EACH SUBFOLDER, IDENTIFIES THE FILES PRESENT, AND MOVES THEM TO A DESIGNATED DESTINATION FOLDER.

Sub sMoveAllFiles()
    Dim fso
    Dim ws As Worksheet
    Dim sRootFolderName As String, dFolderPath As String, FilePattern As String, sSourceFolderName As String
    Dim FindFolder As Variant
    Dim sFolderName As String
    Dim AnswerYes As String
    Dim lastRow As Integer
    
    Dim y As Integer
    Dim dict_pathBlock As Object
    Dim arr_pathBlock As Variant
    
    'CREATE A DICTIONARY TO HOLD BLOCKED FOLDER NAMES
    arr_pathBlock = Array("_Bulk Exports", "_Bulk Exports QA", "GTA Fund Analytical Project", "CRC_Non-Audit Offshore", "TaxOffshore", "Canatfsr10")
    Set dict_pathBlock = CreateObject("Scripting.Dictionary")
    For y = LBound(arr_pathBlock) To UBound(arr_pathBlock) Step 1
        If dict_pathBlock.Exists(arr_pathBlock(y)) Then
        Else
            dict_pathBlock.Add arr_pathBlock(y), ""
        End If
    Next y
    
    Set ws = ThisWorkbook.Worksheets("Move Bulk File")
    Set fso = CreateObject("Scripting.FileSystemObject")
    With ws
        sSourceFolderName = .Cells(1, 2).Value
        dFolderPath = .Cells(2, 2).Value
        FilePattern = "*.xlsx"
        .Cells(5, 2).Value = 0
        .Cells(6, 2).Value = 0
        lastRow = .Cells(Rows.Count, 2).End(xlUp).Row
        .Range("A9:B" & lastRow + 10).Clear
    End With
    
    If Right(sSourceFolderName, 1) = "\" Then
        sSourceFolderName = Left(sSourceFolderName, Len(sSourceFolderName) - 1)
    End If
    
    'VALIDATE THE SOURCE FOLDER PATH
    FindFolder = Split(sSourceFolderName, "\")

    sFolderName = FindFolder(UBound(FindFolder))
    If Replace(sSourceFolderName, sFolderName, "") <> "\\Path\" Then
        MsgBox "This function shall only be use for BulkExport!", vbCritical
        Exit Sub
    End If
    
    If dict_pathBlock.Exists(sFolderName) Then
        MsgBox "Incorrect Folder Selected {" & sFolderName & "}. Please check the Source Folder Path"
        Exit Sub
    End If
    

    If Right(dFolderPath, 1) = "\" Then
        dFolderPath = Left(dFolderPath, Len(dFolderPath) - 1)
    End If
    
    'CHECK IF SOURCE AND DESINTATION FOLDERS EXIST
    If Not fso.FolderExists(sSourceFolderName) Then
        MsgBox "The source file path cannot be found. Please verify the source file path."
        Exit Sub
    End If
    If Not fso.FolderExists(dFolderPath) Then
        MsgBox "The destination file path cannot be found. Please verify the destination file path "
        Exit Sub
    End If
    
    'CHECK IF DESTINATION FOLDER ALREADY CONTAINS FILES
    Dim sFolder As Object: Set sFolder = fso.GetFolder(dFolderPath)
    If sFolder.Files.Count > 0 Then
        AnswerYes = MsgBox("The destination file path already has some files in it. Would you like to proceed?? ", vbQuestion + vbYesNo, "User Response")
        If AnswerYes = vbYes Then
        Else
        Exit Sub
        End If
    End If
    
    'DISABLE SCREEN UPDATING FOR PERFORMANCE
    Application.ScreenUpdating = False
    
    'CALL SUB PROCEDURE TO LIST ALL FOLDERS & SUBFOLDERS
    sbListAllFolders sSourceFolderName & "\", dFolderPath & "\", FilePattern
    
    'CONFIRM DELECTION OF ILES AND SUBFOLDERS IF NECESSARY
    With ws
        If .Cells(5, 2).Value > 0 And .Cells(6, 2).Value = 0 Then
            AnswerYes = MsgBox("Are you certain you want to delete all files and subfolders inside of folder {" & sFolderName & "}? Please be aware that any subfolders and files within it cannot be retrieved once deleted.", vbQuestion + vbYesNo, "User Response")
            If AnswerYes = vbNo Then
                Exit Sub
            End If
            Application.Wait (Now + TimeValue("0:00:05"))
           If fso.FolderExists(sSourceFolderName) Then
            fso.DeleteFolder sSourceFolderName
           End If
           Application.Wait (Now + TimeValue("0:00:05"))
           If Not fso.FolderExists(sSourceFolderName) Then
           fso.CreateFolder sSourceFolderName
           End If
        End If
    End With
    
    'Enable Screen Update
    Application.ScreenUpdating = True
    MsgBox "Completed"
    
End Sub
Sub sbListAllFolders(ByVal sourceFolder As String, ByVal dFolderPath As String, _
        Optional ByVal FilePattern As String)
    
    'Variable Declaration
    Dim oFSO As Object, oSourceFolder As Object, oSubFolder As Object
    Dim iLstRow As Integer
            
    'Create object to FileSystemObject
    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oSourceFolder = oFSO.GetFolder(sourceFolder)

    MoveFiles oSourceFolder.Path, dFolderPath, FilePattern
    'Loop through all Sub folders
    For Each oSubFolder In oSourceFolder.SubFolders
        sbListAllFolders oSubFolder.Path, dFolderPath, FilePattern
    Next oSubFolder

    'Release Objects
    Set oSubFolder = Nothing
    Set oSourceFolder = Nothing
    Set oFSO = Nothing
End Sub
Sub MoveFiles(ByVal SourceFolderPath As String, ByVal DestinationFolderPath As String, Optional ByVal FilePattern As String)
    
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim apSep As String: apSep = Application.PathSeparator
    
    Dim sPath As String: sPath = SourceFolderPath
    If Left(sPath, 1) <> apSep Then sPath = sPath & apSep
        
    Dim sFolder As Object: Set sFolder = fso.GetFolder(sPath)
    
    If sFolder.Files.Count = 0 Then
        Exit Sub
    End If
    
    Dim dPath As String: dPath = DestinationFolderPath
    If Left(dPath, 1) <> apSep Then dPath = dPath & apSep
        
    Dim dFolder As Object: Set dFolder = fso.GetFolder(dPath)
    
    Dim Dict As Object: Set Dict = CreateObject("Scripting.Dictionary")
    Dict.CompareMode = vbTextCompare
    
    Dim sFile As Object
    Dim dFilePath As String
    Dim ErrNum As Long
    Dim MovedCount As Long
    Dim NotMovedCount As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim fileName As String
    Set ws = ThisWorkbook.Worksheets("Move Bulk File")
    With ws
        MovedCount = .Cells(5, 2).Value
        NotMovedCount = .Cells(6, 2).Value
        lastRow = .Cells(Rows.Count, 1).End(xlUp).Row
    End With
    
    For Each sFile In sFolder.Files
        dFilePath = dPath & sFile.Name
        fileName = sFile.Name
        If fso.FileExists(dFilePath) Then
            Dict(sFile.Path) = Empty
            NotMovedCount = NotMovedCount + 1
        Else
            On Error Resume Next
                fso.MoveFile sFile.Path, dFilePath
                ErrNum = Err.Number
                ' e.g. 'Run-time error '70': Permission denied' e.g.
                ' when the file is open in Excel
            On Error GoTo 0
            If ErrNum = 0 Then
                MovedCount = MovedCount + 1
                With ws
                If Right(fileName, 5) = ".xlsx" Then
                lastRow = lastRow + 1
                .Cells(lastRow, 2).Value = fileName
                .Cells(lastRow, 1).Value = Trim(Split(fileName, " - ")(1))
                End If
                End With
            Else
                Dict(sFile.Path) = Empty
                NotMovedCount = NotMovedCount + 1
            End If
        End If
    Next sFile
    
    With ws
        .Cells(5, 2).Value = MovedCount
        .Cells(6, 2).Value = NotMovedCount
    End With
End Sub



