' Needs rework
' Error found:
' Fail to get files to conduct operations

Private folderPath As String
Private yesterdaysWB As Workbook
Private todaysWB As Workbook
Private today As Date

Private yesterdaysWS As Worksheet
Private todaysWS As Worksheet

Public Sub Initialize(folderPath As String)
    today = Date
    folderPath = folderPath
End Sub

Public Sub TransferNotes()
    Call GetXLSXFiles()
End Sub

Private Sub GetXLSXFiles()
    Dim fso As Object
    Dim folder As Object
    Dim fileCount As Integer
    Dim fileNames() As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    fileCount = folder.Files.Count

    If fileCount = 0 Then
        Exit Sub
    End If

    ReDim fileNames(1 To fileCount)

    Dim i As Integer
    i = 1
    For Each file In folder.Files
        fileNames(i) = File.Name
    Next file

    For Each fileName In fileNames
        Dim dateSection As String

        ' slice the string DD_MM_YYYYReport.xlsx to DD_MM
        dateSection = Left(fileName, 5)

        ' add_year(today) to DD_MM
        dateSection = dateSection & "_" & Year(today)

        ' Replace _ with - in DD_MM_Year(today)
        
    Next fileName
End Sub

Private Sub GetXLSXFile(fileName As String) As Workbook

End Sub

Private Sub GetWorkSheets()

End Sub

Private Sub CloseWorkbooks()

End Sub
