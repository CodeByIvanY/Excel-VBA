
With Application.FileDialog(msoFileDialogFilePicker)
     .Title = "Select a File"

     'Set the initial path to the C:\ drive. 
     .InitialFileName = "C:\"

     'Change the text on the button. 
     .ButtonName = "Archive" 

     'Empty the list by clearing the FileDialogFilters collection. 
     .Filters.Clear

     'Add a filter that includes all excel files. 
     .Filters.Add "Excel files","*.xlsx; *.csv; *xls; *xlsm", 1
     
     'Use the Show method to display the File Picker dialog box and return the user's action. 
     If .Show = -1 Then
          FilePath = .SelectedItems(0)
     Else
          MsgBox "No File Selected"
          Exit Sub
     End If

End with

'msoFileDialogFilePicker	3	File Picker dialog box.
'msoFileDialogFolderPicker	4	Folder Picker dialog box.
'msoFileDialogOpen	1	Open dialog box.
'msoFileDialogSaveAs	2	Save As dialog box.

 
With Application.FileDialog(msoFileDialogOpen) 
     'Allow the selection of multiple file. 
     .AllowMultiSelect = True
     .Show 
     ' Display paths of each file selected 
     For lngCount = 1 To .SelectedItems.Count 
          MsgBox .SelectedItems(lngCount) 
     Next lngCount 
End With 

Dim FSO As Object: Set FSO = CreateObject("Scripting.FileSystemObject")
FileCount = FSO.GetFolder(iFolderPath).Files.Count


