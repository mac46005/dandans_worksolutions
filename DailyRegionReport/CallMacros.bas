Public Sub RunHighlightDailyReport()
    Dim hdr As HighlightDailyReport
    Set hdr = New HighlightDailyReport

    hdr.HighlightReport()
End Sub

Public Sub RunTransferNotes()
    Dim tn As TransferNotes
    Set tn = New TransferNotes
    
    tn.Initialize("X:\PROGRAMMING\EXCEL\dandans_worksolutions\reports")
    tn.Execute
End Sub