
With Application.FileDialog(msoFileDialogFilePicker)
     .Title = "Select File"
     If .Show = -1 Then
          FilePath = .SelectedItems(0)
     Else

     End If

End with

'msoFileDialogFilePicker	3	File Picker dialog box.
'msoFileDialogFolderPicker	4	Folder Picker dialog box.
'msoFileDialogOpen	1	Open dialog box.
'msoFileDialogSaveAs	2	Save As dialog box.

Sub UseFileDialogOpen() 
 
    Dim lngCount As Long 
    ' Open the file dialog 
    With Application.FileDialog(msoFileDialogOpen) 
        .AllowMultiSelect = True 
        .Show 
        ' Display paths of each file selected 
        For lngCount = 1 To .SelectedItems.Count 
            MsgBox .SelectedItems(lngCount) 
        Next lngCount 
    End With 
 
End Sub
