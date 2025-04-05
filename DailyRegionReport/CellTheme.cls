' A class that represents a cells theme such as font color and background color. Add more as needed. 



Private background As Long
Private foreground As Long

Public Sub SetBackground(rgb As Long)
    background = rgb
End Sub

Public Function GetBackground() As Long
    GetBackground = background
End Function

Public Sub SetForeground(rgb As Long)
    foreground = rgb
End Sub

Public Function GetForeground() As Long
    GetForeground = foreground
End Function
