Sub ImportData()
    '   Routine to copy data from an excel readable file to a specified worksheet. File format used here is .xls and all related file types. Interface with the dialogue box

    Dim FileToOpen As Variant   'A placeholder for directory accessibility dialogue
    Dim OpenBook As Workbook    'File to import the data from
    
    
    Application.ScreenUpdating = False  'Prevent screen response to running code taking time
    FileToOpen = Application.GetOpenFilename(Title:="Select a File to Import", FileFilter:="Excel Files (*.xls*),*xls*")
    
    '   Open the workbook, copy the data, paste the data in selected sheet,and close the workbook
    '   Paste special ensures that only the values are pasted without formatting or formulas. This can be changed.
    
    If FileToOpen <> False Then
        Set OpenBook = Application.Workbooks.Open(FileToOpen)
        OpenBook.Sheets(1).Range("A:Z").Copy
        ThisWorkbook.Worksheets("Imported File").Range("A1").PasteSpecial xlPasteValues
        OpenBook.Close False
        
        Worksheets("Data").Select
        ActiveSheet.Shapes("FileNameBox").TextFrame.Characters.Text = FileToOpen
        
    End If
    Application.ScreenUpdating = True   'Code may be removed if you feel like because this can be turned on automatically after code is run....but......for precaution sake, I'll leave it here.
End Sub