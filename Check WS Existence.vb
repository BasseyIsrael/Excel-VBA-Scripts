Function xWsExist( _
    WsName As String) As Boolean
   
    '   Function to check the existence of a worksheet in a workbook

    Dim xWs As Worksheet

    xWsExist = False
       
    For Each xWs In Worksheets
        If xWs.Name = xWsName Then
            xWsExist = True
            Exit Function
        End If
    Next xWs
    
End Function
