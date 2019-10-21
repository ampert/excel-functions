Function getFilePath(prevFilePath As String) As String

    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select the file"
        .Filters.Clear
        .Filters.Add "All Files", "*.*"
        .Show
        
        If .SelectedItems.Count <= 0 Then
            getFilePath = prevFilePath
        Else
            getFilePath = .SelectedItems(1)
        End If
    End With
    
End Function

Function getFolderPath(prevPath As String) As String
    
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .Title = "Select folder"
        .AllowMultiSelect = False
        .Show
        
        If .SelectedItems.Count <= 0 Then
            getFolderPath = prevPath
        Else
            getFolderPath = .SelectedItems(1)
        End If
        
    End With

End Function
