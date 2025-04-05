

Public workCompleteTheme As CellTheme
Public passedDueDateAndNotCompleteTheme As CellTheme
Public dueTodayAndNotCompleteTheme As CellTheme
Public passedThreeDaysAndNotCompleteTheme As CellTheme
Public withinThreeDaysAndNotCompleteTheme As CellTheme
Public noSLAAndNoCompleteByTheme As CellTheme

Public Sub Initialize()
    Set workCompleteTheme = new CellTheme
    Call workCompleteTheme.SetForeground(rgb(0,0,0))
    Call workCompleteTheme.SetBackground(rgb(150,193,233))

    Set passedDueDateAndNotCompleteTheme = New CellTheme
    Call passedDueDateAndNotCompleteTheme.SetForeground(rgb(0,0,0))
    Call passedDueDateAndNotCompleteTheme.Background(rgb(192, 0,0))

    Set dueTodayAndNotCompleteTheme = New CellTheme
    Call dueTodayAndNotCompleteTheme.SetForeground(rgb(0,0,0))
    Call dueTodayAndNotCompleteTheme.SetBackground(rgb(242, 133, 239))

    Set passedThreeDaysAndNotCompleteTheme = New CellTheme
    Call passedThreeDaysAndNotCompleteTheme.SetForeground(rgb(0,0,0))
    Call passedThreeDaysAndNotCompleteTheme.SetBackground(rgb(162, 255, 176))

    Set withinThreeDaysAndNotCompleteTheme = New CellTheme
    Call withinThreeDaysAndNotCompleteTheme.SetForeground(rgb(0,0,0))
    Call withinThreeDaysAndNotCompleteTheme.SetBackground(rgb(255,220, 113))

    Set noSLAAndNoCompleteByTheme = New CellTheme
    Call noSLAAndNoCompleteByTheme.SetForeground(rgb(0,0,0))
    Call noSLAAndNoCompleteByTheme.SetBackground(rgb(162, 255, 176))
End Sub