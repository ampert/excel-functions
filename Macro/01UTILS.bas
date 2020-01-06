Function getlastRow(shtName As String, columnRef As String) As Integer

    With ThisWorkbook.Sheets(shtName)
            getlastRow = .Range(startCell & .Rows.Count).End(xlUp).Row
    End With

End Function

Function getlastColumn(shtName As String, columnRef As String) As String

    Column = Sheets(shtName).Range(startCell & "1").CurrentRegion.Columns.Count
    getlastColumn = Split(Cells(1, Column).Address, "$")(1)

End Function

Function openExcelFile() as Workbook

    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    Dim filePath as String
    filePath = fd.SelectedItems(1)

    openExcelFile = Workbooks.Open Filename:=filePath

End Function


Sub CopyPaste_Ex2()
    Sheets("Source").Range("A1:E10").Copy Destination:=Sheets("Destination").Range("A1")
End Sub

Sub copyValues(srcSht As String, srcCol As String, destSht As String, destCol As String)

    Dim lastRow As Integer: lastRow = getlastRow(srcSht, srcCol)
    Dim n_srcCol As String: n_srcCol = srcCol & "2:" & srcCol & lastRow
    Dim n_destCol As String: n_destCol = destCol & "2:" & destCol & lastRow
    ThisWorkbook.Sheets(srcSht).Range(n_destCol).Value = ThisWorkbook.Sheets(destSht).Range(n_srcCol).Value

End Sub

Function columnToInt(columnLetter As String) As Integer

    columnToInt = Range(columnLetter & 1).column

End Function

Sub filter(shtName As String, filterCol As String, filterVal As String)

    ThisWorkbook.Sheets(shtName).Range(filterCol & "1").AutoFilter Field:=columnToInt(filterCol), Criteria1:=filterVal

End Sub

Sub unfilter(shtName As String, filterCol As String)

    ThisWorkbook.Sheets(shtName).Range(filterCol & "1").AutoFilter

End Sub

Sub clearContents(shtName As String, clearCol As String)

    Dim n_clearCol As String: n_clearCol = clearCol & "2"
    ThisWorkbook.Sheets(shtName).Range(n_clearCol, Range(n_clearCol).End(xlDown)).SpecialCells(xlCellTypeVisible).clearContents

End Sub

