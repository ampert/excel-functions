Attribute VB_Name = "excelFunctions"

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

Sub ShowAllSheets()
    
    Dim xWs As Worksheet
    
    For Each xWs In Application.ActiveWorkbook.Worksheets
        xWs.Visible = xlSheetVisible
    Next
    
End Sub

Sub optimizeStart(screen As Boolean, calculation As Boolean, events As Boolean)
    
    If screen Then
        Application.ScreenUpdating = False
    End If
    
    If calculation Then
        Application.calculation = xlCalculationManual
    End If
    
    If events Then
        Application.EnableEvents = False
    End If
    
End Sub

Sub optimizeEnd()

    Application.ScreenUpdating = True
    Application.calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub

Sub CopyPaste_Ex2()
    Sheets("Source").Range("A1:E10").Copy Destination:=Sheets("Destination").Range("A1")
End Sub

Sub createNewSheet(shtName As String, startSheet As Boolean)
    Dim xWs As Worksheet
    
    If startSheet Then
        Set xWs = Sheets.Add(Before:=Sheets(1))
    Else
        Set xWs = Sheets.Add(After:=Sheets(Worksheets.Count))
    End If
    
    xWs.Name = shtName
    
End Sub

Function lastRow(shtName As String, startCell As String) As Integer

    lastRow = Sheets(shtName).Range(startCell).CurrentRegion.Rows.Count

End Function

Function lastColumn(shtName As String, startCell As String) As Integer
    
    lastColumn = Sheets(shtName).Range(startCell).CurrentRegion.Columns.Count

End Function

Function ELOOKUP(lookup_value As Range, lookup_range As Range, value_range As Range) As Variant

    ELOOKUP = Application.WorksheetFunction.Index(value_range, _
        Application.WorksheetFunction.Match(lookup_value, lookup_range, 0))

End Function
