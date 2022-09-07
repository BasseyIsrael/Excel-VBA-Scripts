Sub RectangleRoundedCorners39_Click()
    
    Dim FileToOpen As Variant
    Dim OpenBook As Workbook
    Application.ScreenUpdating = False
    FileToOpen = Application.GetOpenFilename(Title:="Select a File to Import", FileFilter:="Excel Files (*.xls*),*xls*")
    If FileToOpen <> False Then
        Set OpenBook = Application.Workbooks.Open(FileToOpen)
        OpenBook.Sheets(1).Range("A:Z").Copy
        ThisWorkbook.Worksheets("Imported File").Range("A1").PasteSpecial xlPasteValues
        OpenBook.Close False
        
        Worksheets("Data").Select
        ActiveSheet.Shapes("FIleNameBox").TextFrame.Characters.Text = FileToOpen
        
    End If
    Application.ScreenUpdating = True
End Sub