Attribute VB_Name = "ExcelToAccess"
'==============================================================================
'TODO
'==============================================================================
' 1. Stand Alone Module and Initialization
' 2. Run log
' 3. User input for table name



'==============================================================================
'RUN
'==============================================================================

Sub UploadExcelToAccess()
    Call optimize(True)
    
    Dim wbData As Workbook: Set wbData = openExcelFile()
    wbData.Activate
    
    Dim tblName As String: tblName = "Customer_HD"
    
    Call RunQueries(getAccessPath, ExcelToAccessSQL(tblName))
    
    wbData.Close

    MsgBox ("Upload Complete")

    Call optimize(False)
End Sub

'==============================================================================
'MAIN
'==============================================================================

Function ExcelToAccessSQL(tblName As String) As String()

    Dim srcWB As Workbook: Set srcWB = ActiveWorkbook
    Dim srcWS As Worksheet: Set srcWS = srcWB.ActiveSheet
    
    Dim lastColumn As String: lastColumn = getlastColumn("A")
    Dim intlastCol As Integer: intlastCol = srcWS.Range(lastColumn & 1).Column
    Dim lastRow As String: lastRow = getlastRow("A")
    
    'Partition since Excel is limited to 60,000 insert
    Dim part As Integer: part = Application.WorksheetFunction.RoundUp(lastRow / 60000, 0)
    Dim sqlstring() As String
    ReDim sqlstring(part + 3) As String
    
    'Drop / Create Table Scripts
    sqlstring(0) = "DROP TABLE [" & tblName & "]"
    sqlstring(1) = "CREATE TABLE [" & tblName & "] ("
    
    For i = 0 To intlastCol - 1
        With srcWS
            sqlstring(1) = sqlstring(1) & "[" & .Range("A1").Offset(0, i).Value & "] MEMO, "
        End With
    Next i
    
    sqlstring(1) = Left(sqlstring(1), Len(sqlstring(1)) - 2) & ")"
    
    Dim sqlselect As String: sqlselect = "SELECT "
    Dim sqlinsert As String: sqlinsert = "INSERT INTO " & tblName & " "
    Dim sqlfrom As String
    Dim sqlsheet As String
    
    For i = 0 To intlastCol - 1
        With srcWS
            sqlselect = sqlselect & "[" & .Range("A1").Offset(0, i).Value & "], "
        End With
    Next i
    
    sqlselect = Left(sqlselect, Len(sqlselect) - 2) & " "
    sqlfrom = "FROM [Excel 8.0;HDR=YES;DATABASE=" & srcWB.FullName & "]."
    
    If lastRow > 1 And lastRow <= 60000 Then
        
        sqlsheet = "[" & srcWS.Name & "$A1:" & lastColumn & lastRow & "]"
        sqlstring(2) = sqlinsert & sqlselect & sqlfrom & sqlsheet
    
    ElseIf lastRow > 60000 Then
        
        sqlsheet = "[" & srcWS.Name & "$A1:" & lastColumn & 60001 & "]"
        sqlstring(2) = sqlinsert & sqlselect & sqlfrom & sqlsheet
        
        Dim tempWS As Worksheet
        For i = 3 To part + 1
            Set tempWS = srcWB.Sheets.Add
            srcWS.Range("A1:" & lastColumn & "1").Copy tempWS.Range("A1")
            srcWS.Range("A" & (60001 * (i - 1)) + 1 & ":" & lastColumn & (60001 * i)).Copy
            tempWS.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
            Application.CutCopyMode = False
            lastRow = tempWS.Cells(Rows.Count, 1).End(xlUp).Row
            
            sqlsheet = "[" & tempWS.Name & "$A1:" & lastColumn & lastRow & "]"
            sqlstring(i) = sqlinsert & sqlselect & sqlfrom & sqlsheet
        Next i
        
    End If
    
    'Add [UploadDate] to the data
    sqlstring(part + 2) = "ALTER TABLE " & tblName & " ADD COLUMN [UploadDate] DATE"
    sqlstring(part + 3) = "UPDATE " & tblName & " Set [UploadDate] = #" & Format(Now, "mm/dd/yyyy") & "#"
    
    ExcelToAccessSQL = sqlstring
End Function

Sub RunQueries(accessFilePath As String, sqlstring() As String)

    On Error GoTo reason
    Dim scn As String: scn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessFilePath
    Set cn = CreateObject("ADODB.Connection")
    cn.Open scn
    
    For i = 1 To UBound(sqlstring)
        cn.Execute sqlstring(i)
    Next i
    cn.Close
    
    Exit Sub

reason:
    If Err.Number <> -2147217900 Then
        MsgBox (Err.Description)
        End
    End If
    
    For i = 0 To UBound(sqlstring)
        cn.Execute sqlstring(i)
    Next i
    cn.Close

End Sub

'==============================================================================
'FILE DIALOG
'==============================================================================

Function openExcelFile() As Workbook

    MsgBox ("Please select source excel file")
    On Error GoTo reason
    
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd.Filters
        .Clear
        .Add "Excel files", "*.xlsx;*.xls;*.xlsm;*.xlsb"
    End With
    
    fd.Show
    Set openExcelFile = Workbooks.Open(fd.SelectedItems(1), False)
    
    Exit Function
    
reason:
    MsgBox "No file selected"
    End
End Function

Function getAccessPath() As String
    
    MsgBox ("Please select destination access file")
    
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd.Filters
        .Clear
        .Add "Access files", "*.accdb"
    End With
    
    fd.Show
    getAccessPath = fd.SelectedItems(1)

End Function

'==============================================================================
'UTILITIES
'==============================================================================

Function getlastRow(columnRef As String) As Long

    With ActiveWorkbook.ActiveSheet
            x = .Rows.Count
            getlastRow = .Range(columnRef & .Rows.Count).End(xlUp).Row
    End With

End Function

Function getlastColumn(columnRef As String) As String
    
    With ActiveWorkbook.ActiveSheet
        Column = .Range(columnRef & "1").CurrentRegion.Columns.Count
    End With
    getlastColumn = Split(Cells(1, Column).Address, "$")(1)

End Function

Function columnToInt(columnLetter As String) As Long

    columnToInt = Range(columnLetter & 1).Column

End Function

Sub optimize(start As Boolean)

    If start Then
        Application.ScreenUpdating = False
        Application.DisplayStatusBar = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.DisplayAlerts = False
    Else
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.DisplayAlerts = True
    End If

End Sub

'==============================================================================
'LICENSE
'==============================================================================

'MIT License
'
'Copyright (c) 2019 Ampert
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.

