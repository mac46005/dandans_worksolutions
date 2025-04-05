Private ws As Worksheet

Private rowIndex As Long
Private lastRow As Long

Private today As Date

Private themes As DailyReportTheme



Public Sub HighlightReport()
    Initialize

    ' Cycles throw rows A2 to lastRow
    For rowIndex = 2 To lastRow

    Next rowIndex
End Sub

Private Sub Initialize()
    Set ws = ActiveWorkbook.ActiveSheet
    Set themes = New DailyReportTheme
    themes.Initialize
    today = Date
    lastRow = ws.Cells(ws.Rows.Count, DailyRegionReport.IdCol).End(xlUp).Row
End Sub






' Checks to see if current row's STATUS column value has "Work Complete" as its value
Private Function IsWorkComplete() As Boolean
    Dim status As String
    status = LCase(ws.Cells(rowIndex, DailyRegionReport.StatusCol).Value)

    If status = LCase("Work Complete") Then
        IsWorkComplete = True
    Else
        IsWorkComplete = False
    End If
End Function

Private Sub HighlightWorkCompleteRow()
    With GetCurrentRow()
        .Interior.Color = themes.workCompleteTheme.GetBackground()
        .Font.Color = themes.workCompleteTheme.GetForeground()
    End With
End Sub





Private Function IsPassedDueDateAndNotComplete() As Boolean
    If  Then

    Else
        IsPassedDueDateAndNotComplete = False
    End If
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
    Dim completeByCell As Range
    Set completeByCell = ws.Cells(rowIndex, DailyRegionReport.CompleteByCol)

    If IsDate(completeByCell.value) Then
        HasValidCompleteByDate = True
    Else
        HasValidCompleteByDate = False
    End If
End Function

Private Function GetCompleteByDate() As Date
End Function


' Gets the current row using rowIndex from the current for loop cycle in HighlightReport function
Private Function GetCurrentRow() As Range
    Dim rowRange As Range
    Set rowRange = ws.Range(ws.Cells(rowIndex, DailyRegionReport.NotesCol), ws.Cells(rowIndex, DailyRegionReport.LastUpdatedDateCol))
    Set GetCurrentRow = rowRange
End Function
