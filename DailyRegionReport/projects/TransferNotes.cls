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
End Sub

Public Sub GetWorkbooks()
    Dim folder As Object
    Dim file As Object
    Set folder = GetFolder()

    If folder Is Nothing Then Exit Sub
    If folder.Files.Count = 0 Then Exit Sub

    For Each file In folder.Files
        Debug.Print file.Name
    Next file
End Sub

Private Function GetFolder() As Object
    Dim fso As Object
    Dim folder As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Set GetFolder = folder
End Function