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
        If i = (count - 2) Then
            ' TODO get previous xlsx file and set workbook
        Else If i = (count - 1) Then
            ' TODO get new xlsx file and set workbook
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