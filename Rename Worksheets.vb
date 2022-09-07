Sub RenameSheet()
'UpdatebyKutools20191129
Dim xWs As Worksheet
Dim xRngAddress As String
Dim xName As String
Dim xSSh As Worksheet
Dim xInt As Integer
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