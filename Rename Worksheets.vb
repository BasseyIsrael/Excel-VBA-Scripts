Sub RenameWS()

    '   Routine to rename all worksheets in a workbook according to cell values.   

Dim xWs As Worksheet        'For each of the worksheets in the workbook
Dim xRngAddress As String   'Address of the cell that contains the new worksheet name for each worksheet. You can make this an input box
Dim xName As                'Value of the cell from Range
Dim xSSh As Worksheet       'Worksheet Name Value  
Dim xInt As Integer         'Worksheet counter

xRngAddress = Application.ActiveCell.Address
On Error Resume Next

Application.ScreenUpdating = False 'Prevent screen response to running code taking time
For Each xWs In Application.ActiveWorkbook.Sheets
    xName = xWs.Range(xRngAddress).Value
    If xName <> "" Then
        xInt = 0
        Set xSSh = Nothing
        Set xSSh = Worksheets(xName)
        While Not (xSSh Is Nothing)
            Set xSSh = Nothing
            Set xSSh = Worksheets(xName & "(" & xInt & ")")
            xInt = xInt + 1
        Wend
        If xInt = 0 Then
            xWs.Name = xName
        Else
            If xWs.Name <> xName Then
                xWs.Name = xName & "(" & xInt & ")"
            End If
        End If
    End If
Next
Application.ScreenUpdating = True 'Code may be removed if you feel like because this can be turned on automatically after code is run....but......for precaution sake, I'll leave it here.
End Sub