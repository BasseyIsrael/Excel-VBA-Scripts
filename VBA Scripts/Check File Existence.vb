Function ExistCheck( _
    ByVal FilePath As String, _
    ByVal FileName As String, _
    Optional ByVal FileType As String = "NotGiven") As 
    
    '   This function checks a file's existence. The function is called in the routime used to print the result of the existence check.

    Dim fName As String

    ExistCheck = False
    
    FilePath = If(Right(FilePath, 1) <> "\", FilePath & "\", FilePath)
    
    
    '   Extension check to decide on output
    If FileType <> "NotGiven" Then

        If Right(FileName, 1) = "." Then
            FileName = Left(FileName, Len(FileName) - 1)
        End If

        If Left(FileType, 1) <> "." Then
            FileType = "." & FileType
        End If
        
        FullName = FilePath & FileName & FileType
        
    Else
        Fullame = FilePath & FileName
        
    End If
      
    If Dir(fName) <> "" Then
        ExistCheck = True
    End If
        
End Function

Sub FileExistence()
    
    '   A routine for checking if a file exists in a folder location. 
    '   You provide a list of file names and the script goes through the list and checks if they exist in the folderlocation 
    '   If file exists, "Available" is returned, and "Not Available" is returned if it doesnt exist.
    '   NOTE: The status of the file availability can be changes as I have declared them as variables 

    '   This requires the "ExistCheck" Function
    

    '   Define all the variables/options.
    Dim i As Long           'Iteration counter should be able to carry up to one million counts to avoid ocunt limit in excel rows count
    Dim folPath As String     'Folder Location where files should be
    Dim fName As String     'File name to check for (pulled from spreadsheet)
    Dim fType As String     'File type (required by ExistCheck; can be used as array)
    Dim listCol As Long     'Column where file names live
    Dim resCol As Long      'Column where results are printed
    Dim rTrue As String     'Text to return if file exists
    Dim rFalse As String    'Text to return if file does not exist 


    FirstRow = InputBox("Enter first row number", "First Row Number")
    LastRow = InputBox("Enter Last row number", "Last Row Number")
    
    listCol = InputBox("Enter column containing file list", "File list column")
    resCol = nCol+1

    folPath = InputBox("Enter Folder Path", "Folder Path")
    
    resTrue = "Available"
    resFalse = "Not Available"


    '   Loop through row [firstRow] through [lastRow]
    '  take the text in column [listCol] and look for it within [folPath]
    '  If the file exists return [resTrue]; if it does not
    '  return [resFalse] in cell (i, resCol)
    
    For i = FirstRow To LastRow
        fName = Cells(i, listCol)
        
        If ExistCheck(folPath, fName) Then
            Cells(i, resCol) = resTrue
        Else
            Cells(i, resCol) = resFalse
        End If
    Next i


End Sub