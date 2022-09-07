Sub DataReset()

' Keyboard Shortcut: Ctrl+Shift+T
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Hello " & Excel.Application.UserName & ", Are You Sure You Want to Reset Your Data? All running computations will be lost. It is advisable to save your work.", vbYesNo + vbQuestion + vbDefaultButton2, "Reset Data")
    If answer = vbYes Then
        Sheets("Imported File").Select
        Cells.Select
        Selection.ClearContents
        Sheets("Data").Select
        Range("K7").Select
        ActiveSheet.Shapes.Range(Array("FIleNameBox")).Select
        Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = " "
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 1). _
            ParagraphFormat
            .FirstLineIndent = 0
            .Alignment = msoAlignCenter
        End With
        With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 1).Font
            .BaselineOffset = 0
            .NameComplexScript = "+mn-cs"
            .NameFarEast = "+mn-ea"
            .Fill.Visible = msoTrue
            .Fill.ForeColor.ObjectThemeColor = msoThemeColorDark1
            .Fill.ForeColor.TintAndShade = 0
            .Fill.ForeColor.Brightness = 0
            .Fill.Transparency = 0
            .Fill.Solid
            .Size = 9
            .Name = "+mn-lt"
        End With
        Range("A1").Select
    Else
        Exit Sub
    End If
End Sub