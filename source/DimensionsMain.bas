Attribute VB_Name = "DimensionsMain"
Sub Start()
    With New MainView
        .Show vbModeless
    End With
End Sub
Public Function ButtonMovedIn(T)
    With T
        .BackColor = RGB(40, 170, 20)
        .BorderColor = RGB(40, 170, 20)
        .ForeColor = RGB(255, 255, 255)
    End With
End Function
Public Function ButtonMovedOut(T As Label)
    With T
        .BackColor = RGB(240, 240, 240)
        .BorderColor = RGB(100, 100, 100)
        .ForeColor = RGB(0, 0, 0)
    End With
End Function
