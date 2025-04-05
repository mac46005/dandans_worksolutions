Private ws As Worksheet
Private lastRow As Long
Private rowIndex As Long
Private today As Date

Private themes As DailyReportTheme



Public Sub HighlightReport()
End Sub

Private Sub Initialize()
    Set themes = New DailyReportTheme
    themes.Initialize
End Sub

Private Sub SetThemes()
    Set workCompleteTheme = New CellTheme
    Call work
End Sub






Private Function IsWorkComplete() As Boolean

End Function
Private Sub HighlightWorkCompleteRow()
End Sub





Private Function IsPassedDueDateAndNotComplete() As Boolean
End Function

Private Sub HighlightPassedDueDateAndNotCompleteRow()
End Sub







Private Function IsDueDateTodayAndNotComplete() As Boolean

End Function

Private Sub HighlightDueTodayAndNotCompleteRow()

End Sub






Private Function IsPassedThreeDaysAndNotComplete() As Boolean

End Function

Private Sub HighlightPassedThreeDaysAndNotCompleteRow()

End Sub






Private Function IsWithinThreeDaysAndNotComplete() As Boolean
End Function

Private Sub HighlightWithinThreeDaysAndNotCompleteRow()
End Sub



Private Function IsNoSLAAndNoCompleteBy() As Boolean

End Function

Private Sub HighlightNoSLAAndNoCompleteByRow()
End Sub






Private Function GetPriorityColValue() As String
End Function

Private Function HasValidCompleteByDate() As Boolean
End Function

Private Function GetCompleteByDate() As Date
End Function

Private Function GetCurrentRow() As Range
End Function

