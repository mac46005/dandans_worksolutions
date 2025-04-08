Private folderPath As String
Private previousWB As Workbook
Private previousWS As Worksheet
Private newWB As Workbook
Private newWS As Worksheet

Public Sub Initialize(folderPath As String)
    folderPath = folderPath
End Sub

Public Sub Execute()
    GetWorkbooks
    
End Sub

Private Sub GetWorkbooks()
    Dim folder As Object
    folder = GetFolder()

    If folder.Files.Count = 0 Then
        Exit Sub
    End If

    For Each file In folder.Files
        Debug.Print file.Name
    Next file
End Sub
Private Function GetFolder()
    Dim fso As Object
    Dim folder As Object
    Dim fileCount As Integer
    Dim fileNames() As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    set folder = fso.GetFolder(folderPath)
    Set GetFolder = folder
End Function
Private Function GetWorkbook() As Workbook
End Function

Private Sub GetWorksheets()
End Sub

Private Sub CloseWorkbooks()
End Sub

