'This code gets triggered after a save in the workbook

Private Sub Workbook_AfterSave()

End Sub

    Dim conn As New ADODB.Connection
    Dim TableNAme As String
    Dim sqlstr As String
    Dim re As ADODB.Recordset
    
    Dim CRow As Long, TRow As Long
    
    Set conn = New ADODB.Connection
    
    conn.Open "DRIVER={SQL Server}" & ";SERVER=" & "Your Server Name" _
    & ";DATABASE=" & "Your Database Name" _
    & ";UID=" & "Your User ID" _
    & ";PWD=" & "Your Password"
    
    Set rs = New ADODB.Recordset
    
    TRow = Sheets("Sheet Containing Table").Range("A" & Rows.Count).End(xlUp).Row
    
    For CRow = 1 To TRow
        
        Col_1 = Sheets("Sheet Containing Table").Range("A" & CRow).Value
        col_2 = Sheets("Sheet Containing Table").Range("B" & CRow).Value
        col_3 = Sheets("Sheet Containing Table").Range("C" & CRow).Value
        col_4 = Sheets("Sheet Containing Table").Range("D" & CRow).Value
        col_5 = Sheets("Sheet Containing Table").Range("E" & CRow).Value
        col_6 = Sheets("Sheet Containing Table").Range("F" & CRow).Value
        
        sqlstr = "INSERT INTO" & "Table_name" & "VALUES('" & Col_1 & "'," & "'" & col_2 & "'," & "'" & col_3 & "......" ')" 'Ensure you format this line to how you want it to be used or just input an SQL command you want to execute here.
        rs.Open sqlstr, conn, adOpenStatic
        
    Next
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    
    MsgBox ("Upload Complete")
