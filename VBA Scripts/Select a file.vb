Function SelectFile() As String
    '   Function to slecet a file from a dialogue box, not minding the file format.
    
        'Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = False
        .Show
 
        'Set function result to selected filename
        If .SelectedItems.Count <> 0 Then
            SelectFileName = .SelectedItems(1)
        End If
    
    End With
 
End Function