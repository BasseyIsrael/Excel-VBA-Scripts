Sub SaveWorksheetsAsCsv()

'Routine to help you save all Worksheets in a Workbook as stand alone CSV files

Dim WS As Excel.Worksheet
Dim SaveToDirectory As String
Dim CurrentWorkbook As String
Dim CurrentFormat As Long

CurrentWorkbook = ThisWorkbook.FullName
CurrentFormat = ThisWorkbook.FileFormat

' Specify the Directory that you would like your XLSX files saved in
' The file names of the Workbooks is a function of the names of the Worksheets
'   SaveToDirectory = InputBox("Enter Folder Path", "Folder Path")

namePrefix = InputBox("Enter a name prefix for all files", "Name Prefix")

With Application.FileDialog(msoFileDialogFolderPicker)
        
        If .Show = True Then
            DestFolder = .SelectedItems(1)
        Else
            MsgBox "You must specify a folder to save the PDF into." & vbCrLf & vbCrLf & "Press OK to exit.", vbCritical, "Must Specify Destination Folder"    
        End If
        
    End With

For Each WS In ThisWorkbook.Worksheets
    Sheets(WS.Name).Copy
    ActiveWorkbook.SaveAs Filename:=DestFolder & Application.PathSeparator & namePrefix & "-" & WS.Name & ".csv", FileFormat:=xlCSV
    ActiveWorkbook.Close savechanges:=False
    ThisWorkbook.Activate
Next

Application.DisplayAlerts = False
ThisWorkbook.SaveAs Filename:=CurrentWorkbook, FileFormat:=CurrentFormat
Application.DisplayAlerts = True
' Temporarily turn alerts off to prevent the user being prompted
'  about overwriting the original file.

End Sub