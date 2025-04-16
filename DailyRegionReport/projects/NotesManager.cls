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
    ' TODO Transfer Notes from Worksheet Objects
    TransferNotes

End Sub

Public Sub GetWorkbooks()
    Dim folder As Object
    Dim file As Object
    Dim count As Integer

    Set folder = GetFolder()
    count = folder.Files.Count

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
            previousWB = GetWorkbook(file)
        Else If i = (count - 1) Then
            newWB = GetWorkbook(file)
        End If
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
    Dim wb As Workbook
    Set wb = Workbooks.Open(file.Name)
    Set GetWorkbook = wb
End Function







Private Sub SetWorksheets()
    previousWS = previousWB.Sheets(1)
    nextWS = nextWB.Sheets(1)
End Sub






Private Sub TransferNotes()
    Dim newLastRow As Long
    newLastRow = newWS.Cells(newWS.Rows.Count, DailyRegionReport.IdCol).Row
    
    Dim newRange As Range
    Set newRange = newWS.Range(newWS.Cells(2, DailyRegionReport.IdCol), newWS.Cells(newLastRow, DailyRegionReport.IdCol))


    Dim previousLastRow As Long
    previousLastRow = previousWS.Cells(previousWS.Rows.Count, DailyRegionReport.IdCol).Row

    Dim previousRange As Range
    Set previousRange = previousWS.Range(previousWS.Cells(2, DailyRegionReport.IdCol), preciousWS.Cells(preciousLastRow, DailyRegionReport.IdCol))

    For Each nCell In newRange
        For Each pCell In previousRange
            If nCell.Value = pCell.Value
                
            End If
        Next pCell 
    Next nCell
    
End Sub

Private Sub CloseWorkbooks()
    previousWB.Close
End Sub