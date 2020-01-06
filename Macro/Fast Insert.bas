Reference:
Microsoft ActiveX Data Objects 6.1 Library


Map Database
Upload SNOW Data
Upload CCCI Data

Sub UploadSNOWData(
excelSheetName as String,
accessFilePath As String,
accessTableName As String)
    'Create New Dummy Sheet UploadData
    'Open the file, Paste values all data to New Dummy Sheet
    Dim scn as String: scn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessFilePath
    Set cn = CreateObject("ADODB.Connection")
    cn.Open scn
    cn.Execute ssql
    cn.Close
End Sub

'Requirements: Data starts at A1, Columns Row 1, 1 sheet only
Function ExcelToRangeSQL(excelSheetName as String) as String
    Dim lastColumn as String : lastColumn = getlastRow(excelSheetName, "A1")
    Dim lastRow as String : lastRow = getlastColumn(excelSheetName, "A1")
    Dim sqlselect as String
    Dim sqlfrom as String
    sqlfrom = "FROM [Excel 8.0;HDR=YES;DATABASE=" & _
    ThisWorkbook.FullName & "]." & _
    "[" & ThisWorkbook.excelSheetName & _
    "$A1:" & lastColumn & lastRow & "]"
    ExcelToRangeSQL = sqlselect & " " & sqlfrom
End Sub