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
    Debug.Print "GetWorkbooks()"

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

    Dim prevFile As Object
    Dim newFile As Object

    Dim i As Integer
    i = 0
    For Each file In folder.Files
        Debug.Print file.Name
        ' TODO May need to add logic to see whether file has been modified recently or a while.
        If i = (count - 2) Then
            Set prevFile = file
        ElseIf i = (count - 1) Then
            Set newFile = file
        End If
        i = i + 1
    Next file


    Debug.Print "Workbooks got are:\n" & "prev: " & prevFile.Name & "\n" & "next: " & newFile.Name
    Set previousWB = GetWorkbook(prevFile)
    Set newWB = GetWorkbook(newFile)
End Sub

Private Function GetFolder() As Object
    Dim fso As Object
    Dim folder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Set GetFolder = folder
End Function

Private Function GetWorkbook(file As Object) As Workbook
    Dim wb As Workbook

    Set wb = Workbooks.Open(folderPath & "/" & file.Name)
    Set GetWorkbook = wb
End Function







Private Sub SetWorksheets()
    Set previousWS = previousWB.Sheets(1)
    Set nextWS = newWB.Sheets(1)
End Sub


Private Function GetIdColRange(ws As Worksheet) As Range
    Dim lastRow As Long
    Dim idRange As Range

    lastRow = ws.Cells(ws.Rows.Count, DailyRegionReport.IdCol).End(xlUp).Row
    Set idRange = ws.Range(ws.Cells(2, DailyRegionReport.IdCol), ws.Cells(lastRow, DailyRegionReport.IdCol))
    Set GetIdColRange = idRange
End Function




Private Sub TransferNotes()
    Dim newRange As Range
    Dim previousRange As Range
    
    Set newRange = GetIdColRange(newWS)
    Set previousRange = GetIdColRange(previousWS)
    
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


