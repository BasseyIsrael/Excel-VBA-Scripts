Sub SetFull_screen()
    '   This routine helps you activate the excel full screen feature. It best works when connected to a button or a shortcut. 

    '   Check to see if fullscreen is already turned on
    If  ActiveWindow.DisplayWorkbookTabs = False And _
        ActiveWindow.DisplayHorizontalScrollBar = False And _
        ActiveWindow.DisplayVerticalScrollBar = False And _
        Application.DisplayFullScreen = True _
    
    Then
    
        MsgBox "Full Screen already turned on"
    
    Else
        ActiveWindow.DisplayWorkbookTabs = False
        ActiveWindow.DisplayHorizontalScrollBar = False
        ActiveWindow.DisplayVerticalScrollBar = False
        Application.DisplayFullScreen = True
    
    End If
    
End Sub

Sub ExitFull_screen()

    '   This routine helps you deactivate the excel full screen feature. It best works when connected to a button or a shortcut. 

    '   Check to see if fullscreen is already turned off

    If  ActiveWindow.DisplayWorkbookTabs = True And _
        Application.DisplayFullScreen = False And _
        ActiveWindow.DisplayHorizontalScrollBar = True _
    
    Then
    
        MsgBox "Full Screen already turned off"
    
    Else
        ActiveWindow.DisplayWorkbookTabs = True
        Application.DisplayFullScreen = False
        ActiveWindow.DisplayHorizontalScrollBar = True
    
    End If
End Sub
