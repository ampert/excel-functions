References:
OLE Automation
Microsoft ActiveX Data Objects 6.1 Library

Sub RunQuery(accessFilePath As String, sqlQuery As String)

    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
        accessFilePath & _
        ";Persist Security Info=False;"
    cn.Open
    cn.Execute (sqlQuery)

End Sub

Sub ClearTable(accessFilePath As String, tableName As String)
    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
        accessFilePath & _
        ";Persist Security Info=False;"
    cn.Open
    cn.Execute ("DELETE FROM " & tableName)
End Sub

Sub AccesstoExcel(excelSheet As String, accessFilePath As String, sqlQuery As String)

    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
        accessFilePath & _
        ";Persist Security Info=False;"
    cn.Open
    
    Set rst = New ADODB.Recordset
    
    rst.Open Source:=sqlQuery, ActiveConnection:=cn
    
    ThisWorkbook.Sheets(excelSheet).Cells.Delete
    For i = 0 To rst.Fields.Count - 1
        ThisWorkbook.Sheets(excelSheet).Range("A1").Offset(0, i).Value = rst.Fields(i).Name
    Next i
    ThisWorkbook.Sheets(excelSheet).Range("A2").CopyFromRecordset rst
    
End Sub

Sub ExceltoAccess(excelFilePath As String, accessFilePath As String, accessTableName As String, sheetNum As Integer)
    Dim wbRef As Workbook
    Dim arrayRef As Variant
    Dim fieldCount As Long
    
    Set wbRef = Workbooks.Open(excelFilePath)
    arrayRef = wbRef.Sheets(sheetNum).Range("A1").CurrentRegion.Value
    wbRef.Close
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
        accessFilePath & _
        ";Persist Security Info=False;"
    cn.Open
    
    Set rst = New ADODB.Recordset
    rst.Open accessTableName, cn, adOpenKeyset, adLockPessimistic, adCmdTable
    
    ' Get the total numbers of fields in the table
    fieldCount = rst.Fields.Count
    
    For i = LBound(arrayRef, 1) + 1 To UBound(arrayRef, 1)
        rst.AddNew
        rst(1).Value = ThisWorkbook.Sheets(1).Range("B7").Value '[TEST]for different upload date
        'rst(1).Value = Date 'rst(1) is upload date
        For j = 1 To fieldCount - 7
            rst(j + 1).Value = arrayRef(i, j)
        Next j
        rst.Update
    Next i
    
    rst.Close
    cn.Close
    
End Sub

Sub LocalExceltoAccess(accessFilePath As String, accessTableName As String, sheetName As String)
    Dim arrayRef As Variant
    Dim fieldCount As Long
    
    arrayRef = ThisWorkbook.Sheets(sheetName).Range("A1").CurrentRegion.Value
    
    Set cn = New ADODB.Connection
    cn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & _
        accessFilePath & _
        ";Persist Security Info=False;"
    cn.Open
    
    Set rst = New ADODB.Recordset
    rst.Open accessTableName, cn, adOpenKeyset, adLockPessimistic, adCmdTable
    
    ' Get the total numbers of fields in the table
    fieldCount = rst.Fields.Count
    
    For i = LBound(arrayRef, 1) + 1 To UBound(arrayRef, 1)
        rst.AddNew
        For j = 1 To fieldCount - 1
            rst(j).Value = arrayRef(i, j)
        Next j
        rst.Update
    Next i
    
    rst.Close
    cn.Close
    
End Sub