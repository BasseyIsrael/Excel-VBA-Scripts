Sub Full_screen()
    If ActiveWindow.DisplayWorkbookTabs = False And ActiveWindow.DisplayHorizontalScrollBar = False And ActiveWindow.DisplayVerticalScrollBar = False And _
    Application.DisplayFullScreen = True Then
    
    MsgBox "Full Screen already turned on"
    
    Else
    ActiveWindow.DisplayWorkbookTabs = False
    ActiveWindow.DisplayHorizontalScrollBar = False
    ActiveWindow.DisplayVerticalScrollBar = False
    Application.DisplayFullScreen = True
    
    End If
    
End Sub

Sub ExitFull_screen()

    If ActiveWindow.DisplayWorkbookTabs = True And Application.DisplayFullScreen = False And ActiveWindow.DisplayHorizontalScrollBar = True Then
    
    MsgBox "Full Screen already turned off"
    
    Else
    ActiveWindow.DisplayWorkbookTabs = True
    Application.DisplayFullScreen = False
    ActiveWindow.DisplayHorizontalScrollBar = True
    
    End If
End Sub
