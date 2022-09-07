Sub ListPDFFiles()

    '   Routine to obtain a list of all the pdf files contained in a folder, givene the folder path.
     
    Dim objFSO As Object
    Dim objFolder As Object
    Dim objFile As Object
    Dim xWs As Worksheet
    Dim folderpath As String
    
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set xWs = ActiveSheet
    
    
    folderpath = InputBox("Enter Folder Path", "Folder Path")
     
    'Get the folder object associated with the directory
    Set objFolder = objFSO.GetFolder(folderpath)
    xWs.Cells(1, 1).Value = "The files found in " & objFolder.Name & " are:"
       
     'Loop through the Files collection
    For Each objFile In objFolder.Files
        If UCase(Right(objFile.Name, 4)) = ".PDF" Then
            ws.Cells(ws.UsedRange.Rows.Count + 1, 1).Value = Replace$(UCase$(objFile.Name), ".PDF", "")
        Z = Z + 1

        End If
      Next
      If Z = 0 Then

             MsgBox "No PDF Files found"

        End If
    
    'Clean up!
    Set objFolder = Nothing
    Set objFile = Nothing
    Set objFSO = Nothing
    
End Sub