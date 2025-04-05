' Represents the column names for the sheet
Public Enum DailyRegionReport
    NotesCol = 1
    IdCol = 2
    CompleteByCol = 3
    PropertyCol = 4
    TypeCol = 5
    AssignedToCol = 6
    StatusCol = 7
    PriorityCol = 8
    CommentsCol = 9
    LastUpdatedDateCol = 10
End Enum

' Gets a row Range object from given WorkSheet
Public Function GetRow(ws As WorkSheet, rowIndex As Long, endColIndex As Long) As Range
    Dim rowRange As Range
    Set rowRange = ws.Range(ws.Cells(rowRange, 1), ws.Cells(rowIndex, endColIndex))
    Set GetRow = rowRange
End Function