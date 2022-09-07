Sub RenameWS()

    '   Routine to rename all worksheets in a workbook according to cell values.   

Dim xWs As Worksheet        'For each of the worksheets in the workbook
Dim xRngAddress As String   'Address of the cell that contains the new worksheet name for each worksheet. You can make this an input box
Dim xName As                'Value of the cell from Range
Dim xSSh As Worksheet       'Worksheet Name Value  
Dim xInt As Integer         'Worksheet counter

xRngAddress = Application.ActiveCell.Address
On Error Resume Next

Application.ScreenUpdating = False
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
Application.ScreenUpdating = True
End Sub