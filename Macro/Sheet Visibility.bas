Sub ShowAllSheets()
    
    Dim xWs As Worksheet
    
    For Each xWs In Application.ActiveWorkbook.Worksheets
        xWs.Visible = xlSheetVisible
    Next
    
End Sub

Sub HideAllSheetsExcept(shtName As String, veryHidden As Boolean)

    Dim xWs As Worksheet, hideValue As Integer
    
    If veryHidden Then
        hideValue = xlSheetVeryHidden
    Else
        hideValue = xlSheetHidden
    End If
    
    For Each xWs In Application.ActiveWorkbook.Worksheets
        If xWs.Name <> shtName Then
            xWs.Visible = hideValue
        End If
    Next
    
End Sub