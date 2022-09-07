Sub RectangleRoundedCorners3_Click()
    On Error GoTo ErrHandler
    
    ' SET Outlook APPLICATION OBJECT.
    Dim objOutlook As Object
    Set objOutlook = CreateObject("Outlook.Application")
    
    ' CREATE EMAIL OBJECT.
    Dim objEmail As Object
    Set objEmail = objOutlook.CreateItem(0)
    Dim CurrentMonth As String, DestFolder As String, PDFFile As String
    Dim OpenPDFAfterCreating As Boolean, AlwaysOverwritePDF As Boolean, DisplayEmail As Boolean
    Dim OverwritePDF As VbMsgBoxResult
    
    OpenPDFAfterCreating = True
    AlwaysOverwritePDF = False
    
     With Application.FileDialog(msoFileDialogFolderPicker)
        
        If .Show = True Then
        
            DestFolder = .SelectedItems(1)
            
        Else
        
            MsgBox "You must specify a folder to save the PDF into." & vbCrLf & vbCrLf & "Press OK to exit this macro.", vbCritical, "Must Specify Destination Folder"
                
            Exit Sub
            
        End If
        
    End With
    
    CurrentDate = Workbooks("PEGIS GLOBAL SERVICES PERSONNEL CREDENTIALS STATUS.XLSM").Sheets("Data Summary").Range("C2").Value
    PDFFile = DestFolder & Application.PathSeparator & "Staff Credential Update" _
                & "_" & CurrentDate & ".pdf"
    If Len(Dir(PDFFile)) > 0 Then
    
        If AlwaysOverwritePDF = False Then
        
            OverwritePDF = MsgBox(PDFFile & " already exists." & vbCrLf & vbCrLf & "Do you want to overwrite it?", vbYesNo + vbQuestion, "File Exists")
        
            On Error Resume Next
            'If you want to overwrite the file then delete the current one
            If OverwritePDF = vbYes Then
    
                Kill PDFFile
        
            Else
    
                MsgBox "OK then, if you don't overwrite the existing PDF, I can't continue." _
                    & vbCrLf & vbCrLf & "Press OK to exit this macro.", vbCritical, "Exiting Macro"
                
                Exit Sub
        
            End If

        Else
        
            On Error Resume Next
            Kill PDFFile
            
        End If
        
        If Err.Number <> 0 Then
        
            MsgBox "Unable to delete existing file.  Please make sure the file is not open or write protected." _
                    & vbCrLf & vbCrLf & "Press OK to exit this macro.", vbCritical, "Unable to Delete File"
                
            Exit Sub
        
        End If
            
    End If
   

    'Create the PDF
    Workbooks("PEGIS GLOBAL SERVICES PERSONNEL CREDENTIALS STATUS.XLSM").Sheets("Data Summary").ExportAsFixedFormat Type:=xlTypePDF, Filename:=PDFFile, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas _
        :=False, OpenAfterPublish:=OpenPDFAfterCreating
    With objEmail
        .To = "Fadoyeni.tunji@pegisglobal.com"
        .Subject = "Personnel Credentials Status"
        .Body = "Dear Sir, Kindly find the update on employee credentials attached to the email. Regards."
        .Attachments.Add PDFFile
        .Send
    End With
    
    ' CLEAR.
    Set objEmail = Nothing:    Set objOutlook = Nothing
        
ErrHandler:
End Sub