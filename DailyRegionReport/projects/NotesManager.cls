Private folderPath As String
Private previousWB As Workbook
Private previousWS As Worksheet
Private newWB As Workbook
Private newWS As Worksheet

Public Sub Initialize(path As String)
    folderPath = path
End Sub

Public Sub Execute()
    GetWorkbooks
    SetWorksheets
    TransferNotes
    CloseWorkbooks
End Sub

Private Sub GetWorkbooks()
    Dim folder As Object
    Dim file As Object
    Dim count As Integer

    Set folder = GetFolder()
    count = folder.Files.count

    If folder Is Nothing Then Exit Sub
    If count = 0 Then Exit Sub

    ' Loop through Collection of Files
    ' Get the last two files :(
    ' Apparently, microsoft recommends to loop through the Files Object because the Object is of type Collection. The type of Collection given is completely different than the one that is given in Excel VBA.
    Dim i As Integer
    i = 0
    For Each file In folder.Files
        Debug.Print file.Name
        ' TODO May need to add logic to see whether file has been modified recently or a while.
        If i = (count - 2) Then
            Set previousWB = GetWorkbook(file)
        ElseIf i = (count - 1) Then
            Set newWB = GetWorkbook(file)
        End If
        i = i + 1
    Next file
End Sub

Private Function GetFolder() As Object
    Dim fso As Object
    Dim folder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Set GetFolder = folder
End Function

Private Function GetWorkbook(file As Object) As Workbook

    ' HAVING ISSUE HERE
    ' file being grab starts with ~somefile.xlsx
    ' but i dont want that. I want the next and previous files.

    Dim wb As Workbook
    Set wb = Workbooks.Open(folderPath & "/" & file.Name)
    Set GetWorkbook = wb
End Function







Private Sub SetWorksheets()
    Set previousWS = previousWB.Sheets(1)
    Set nextWS = nextWB.Sheets(1)
End Sub




Private Function GetPreviousRange() As Range
    Dim previousRange As Range
    Dim previousLastRow As Long

    previousLastRow = previousWS.Cells(preciousWS.Rows.count, DailyRegionReport.IdCol).End(xlUp).Row
    Set previousRange = preciousWS.Range(preciousWS.Cells(2, DailyRegionReport.IdCol), preciousWS.Cells(previousLastRow, DailyRegionReport.IdCol))
    Set GetPreviousRange = previousRange
End Function


Private Function GetNewRange() As Range
    Dim newRange As Range
    Dim newLastRow As Long

    newLastRow = newWS.Cells(newWS.Rows.count, DailyRegionReport.IdCol).End(xlUp).Row
    Set newRange = newWS.Range(newWS.Cells(2, DailyRegionReport.IdCol), newWS.Cells(newLastRow, DailyRegionReport.IdCol))
    Set GetNewRange = newRange
End Function

Private Sub TransferNotes()
    Dim newLastRow As Long
    newLastRow = newWS.Cells(newWS.Rows.count, DailyRegionReport.IdCol).Row
    
    Dim newRange As Range
    Set newRange = newWS.Range(newWS.Cells(2, DailyRegionReport.IdCol), newWS.Cells(newLastRow, DailyRegionReport.IdCol))


    Dim previousLastRow As Long
    previousLastRow = previousWS.Cells(previousWS.Rows.count, DailyRegionReport.IdCol).End(xlUp).Row

    Dim previousRange As Range
    Set previousRange = previousWS.Range(previousWS.Cells(2, DailyRegionReport.IdCol), preciousWS.Cells(preciousLastRow, DailyRegionReport.IdCol))

    For Each nCell In newRange

        For Each pCell In previousRange
            If nCell.value = pCell.value Then
                ' Transfer note

                Dim nNoteCell As Range
                Set nNoteCell = newWS.Cells(nCell.Row, DailyRegionReport.NotesCol)
                nNoteCell.value = previousWS.Cells(pCell.Row, DailyRegionReport.NotesCol).value

                Exit For
            End If
        Next pCell
    Next nCell
    
End Sub

Private Sub CloseWorkbooks()
    previousWB.Close
End Sub
