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
