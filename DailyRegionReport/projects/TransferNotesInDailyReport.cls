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
        dateSection = dateSection & "_" & Year(today)

        ' Replace _ with - in DD_MM_Year(today)
        dateSection = Replace(dateSection, "_", "-")

        ' Create date Object
        Dim fileDate As Date

        ' Initialize date Object with dateSection value
        fileDate = dateSection

        ' Check if file date is equal to today and if file date is yesterday
        If Int(fileDate) - Int(today) = 0 Then ' Get yesterdays XLSX file
            Set todaysWB = GetXLSXFile(fileName)
        ElseIf Int(fileDate) - Int(today) = -1 Then ' Get today's XLSX file
            Set yesterdaysWB = GetXLSXFile(fileName)
        End If

        ' throw error if today or yesterdays files are nothing
        ' If todaysWB Is Not Nothing Or yesterdaysWB Is Not Nothing Then
        If todaysWB Is Not Nothing Or yesterdaysWB Is Not Nothing Then
            Exit For
        End IF
        
    Next fileName

    MsgBox "Files have been successfully!",,"File Transfer Successful" 
End Sub

Private Sub GetXLSXFile(fileName As String) As Workbook
    
End Sub

Private Sub GetWorkSheets()

End Sub

Private Sub CloseWorkbooks()

End Sub
