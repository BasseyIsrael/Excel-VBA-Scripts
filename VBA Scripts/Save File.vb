Function SaveFileAs( _
  FileName as String
  FilePath as string
  Optional WorkbookName as Workbook = ActiveWorkbook.Name) as Boolean
  
  ' Function to save file as an excel file with "xlsx" file format
  
  
  Workbooks(WorkBookName).SaveAs _
    FileName:=FileName
    
    ActiveWorkbook.SaveAs Filename:=thisWb.Path & "\new workbook.xlsx"
    ActiveWorkbook.Close savechanges:=False
End Function