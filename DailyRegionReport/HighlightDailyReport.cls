Private ws As Worksheet

Private rowIndex As Long
Private lastRow As Long

Private today As Date

Private themes As DailyReportThemes



Public Sub HighlightReport()
    Initialize

    ' Cycles throw rows A2 To lastRow
    For rowIndex = 2 To lastRow

    Next rowIndex
End Sub

Private Sub Initialize()
    Set ws = ActiveWorkbook.ActiveSheet
    Set themes = New DailyReportThemes
    themes.Initialize
    today = Date
    lastRow = ws.Cells(ws.Rows.Count, DailyRegionReport.IdCol).End(xlUp).Row
End Sub






' Checks To see If current row's STATUS column value has "Work Complete" As its value
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
    If HasValidCompleteByDate() Then
        Dim completeByValue As Date
        completeByValue = GetCompleteByDate()

        If completeByValue < today Then ' If row's complete by value is less than today
            If Not IsWorkComplete() Then
                IsPassedDueDateAndNotComplete = True
             Exit Function
            End If
        End If
        IsPassedDueDateAndNotComplete = False
End Function

Private Sub HighlightPassedDueDateAndNotCompleteRow()
    With GetCurrentRow()
        .Interior.Color = passedDueDateAndNotCompleteTheme.GetBackground()
        .Font.Color = passedDueDateAndNotCompleteTheme.GetForeground()
    End With
End Sub









Private Function IsDueDateTodayAndNotComplete() As Boolean
    If HasValidCompleteByDate() Then
        If IsCompleteByToday() And Not IsWorkComplete() Then
            IsDueTodayAndNotComplete = True
         Exit Function
        End If
    End If

    IsDueDateTodayAndNotComplete = False

End Function

Private Sub HighlightDueTodayAndNotCompleteRow()
    With GetCurrentRow()
        .Interior.Color = themes.passedDueDateAndNotCompleteTheme.GetBackground()
        .Font.Color = themes.passedDueDateAndNotCompleteTheme.GetForeground()
    End With
End Sub






Private Function IsPassedThreeDaysAndNotComplete() As Boolean
    If HasValidCompleteByDate() Then
        Dim completeByValue As Date
        completeByValue = GetCompleteByDate()

        Dim dateDifference As Integer
        dateDifference = (Int(today) - Int(completeByValue))

        If dateDifference <= -3 Then
            If Not IsWorkComplete() Then
                IsPassedThreeDaysAndNotComplete = True
                Exit Function
            End If
        End If
    End If

    IsPassedThreeDaysAndNotComplete = False
End Function

Private Sub HighlightPassedThreeDaysAndNotCompleteRow()
    With GetCurrentRow()
        .Interior.Color = themes.passedThreeDaysAndNotCompleteTheme.GetBackground()
        .Font.Color = themes.passedThreeDaysAndNotCompleteTheme.GetForeground()
    End With
End Sub






Private Function IsWithinThreeDaysAndNotComplete() As Boolean
    If HasValidCompleteByDate() Then
        Dim completeByValue As Date
        completeByValue = GetCompleteByDate()

        If Not IsWorkComplete() And DifferenceFromToday(completeByValue) <= -1 And DifferenceFromToday(completeByValue) >= -3 Then
            IsWithinThreeDaysAndNotComplete = True
            Exit Function
        End If
    End If

    IsWithinThreeDaysAndNotComplete = False
End Function

Private Sub HighlightWithinThreeDaysAndNotCompleteRow()
    With GetCurrentRow()
        .Interior.Color = themes.withinThreeDaysAndNotCompleteTheme.GetBackground()
        .Font.Color = themes.withinThreeDaysAndNotCompleteTheme.GetForeground()
    End With
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



' Get the current rowIndex's 'Complete By' column value As a Date object
Private Function GetCompleteByDate() As Date
    Dim completeByValue As Date
    completeByValue = wsCells(rowIndex, DailyRegionReport.CompleteByCol).value

    GetCompleteByDate = completeByValue
End Function


' Gets the current row using rowIndex from the current For Loop cycle in HighlightReport Function
Private Function GetCurrentRow() As Range
    Dim rowRange As Range
    Set rowRange = ws.Range(ws.Cells(rowIndex, DailyRegionReport.NotesCol), ws.Cells(rowIndex, DailyRegionReport.LastUpdatedDateCol))
    Set GetCurrentRow = rowRange
End Function


Private Function IsCompleteByToday() As Boolean
    Dim completeByValue As Date
    completeByValue = GetCompleteByDate()

    If (Int(completeByValue) = Int(today)) Then
        IsCompleteByToday = True
    Else
        IsCompleteByToday = False
    End If
End Function

Private Function DifferenceFromToday(d As Date) As Int
    Dim result As Int
    result = (Int(today) - Int(d))

    DifferenceFromToday = result
End Function