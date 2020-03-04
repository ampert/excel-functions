'==============================================================================
'EXCEL ACCESS LOADER
'==============================================================================
'Take any excel file and load it to access with just a few clicks!

'==============================================================================
'REQUIREMENTS AND LIMITATIONS
'==============================================================================
' 1. Excel file must have the sheet with the data active upon open
' 2. Column names must be in the first row
' 3. Column names should not be duplicated
' 4. No cell must be merged
' 5. First column should not contain blanks

'==============================================================================
'INSTALLATION GUIDE
'==============================================================================
'Import this module to any excel file and add the following references below:
'   - Microsoft Office 16.0 Object Library
'Edit constants for your requirements
'Run Install()

'==============================================================================
'CONSTANTS
'==============================================================================
Const startReportDate = "1/1/2020"
Const endReportDate = "1/1/2050"

Const tableList = _
"TABLE_1" & "," & _
"TABLE_2" & "," & _
"TABLE_3" & "," & _
"TABLE_4"

'==============================================================================
'VARIABLES
'==============================================================================

Dim accessPath As String
Dim reportDate As Date
Dim tblName As String

Sub VarInit()
    accessPath = ThisWorkbook.Sheets(1).Range("D3")
    reportDate = ThisWorkbook.Sheets(1).Range("D7")
    tblName = ThisWorkbook.Sheets(1).Range("D9")
End Sub

'==============================================================================
'INSTALL
'==============================================================================

Sub Install()
    choice = MsgBox("All data on this excel file will be cleared. Do you wish to continue?" _
        , vbYesNo + vbExclamation, "Install Excel Access Loader")
    
    If choice = vbNo Then
        Exit Sub
    End If
    
    Call optimize(True)
    
    Call DeleteAllSheets
    
    With ThisWorkbook.Sheets(1)
        .Columns("C:C").ColumnWidth = 1.5
        .Columns("B:B").ColumnWidth = 12.8
        .Columns("D:D").ColumnWidth = 12.8
        .Cells.Interior.Color = vbWhite
        
        .Range("B3:B5").Merge
        .Range("B3").Value = "Database Path"
        .Range("B7").Value = "Reporting Date"
        .Range("B9").Value = "Table Name"
        
        'Title Box Color, Border, Alignment
        With .Range("B3:B5,B7,B9")
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = RGB(95, 155, 213)
            .Font.Color = RGB(255, 255, 255)
        End With
        
        'Data Box Border and Alignment
        .Range("D3:I5").Merge
        With .Range("D3:I5,D7,D9")
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
        End With
        
        'Report Date Validation
        .Range("D7").Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=startReportDate, Formula2:=endReportDate
        
        'Table Name Validation
        .Range("D9").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=tableList
        
        'Map Button
        .Buttons.Add(460, 35, 120, 30.5).Select
        Selection.Name = "Map_DB"
        Selection.Characters.Text = "Map Database"
        Selection.OnAction = "getAccessPath"
        
        'Upload Button
        .Buttons.Add(160, 150, 120, 30.5).Select
        Selection.Name = "Upload"
        Selection.Characters.Text = "Upload To Database"
        Selection.OnAction = "UploadExcelToAccess"
        
        .Range("A1").Select
    End With
    
    Call optimize(False)
    
    MsgBox "Installation Complete"
End Sub

'==============================================================================
'BUTTONS
'==============================================================================

Sub UploadExcelToAccess()
    Call optimize(True)
    
    Call VarInit
    
    If reportDate = 0 Then
        MsgBox "Report Date is required"
        GoTo blank
    ElseIf tblName = "" Then
        MsgBox "Table Name is required"
        GoTo blank
    End If
    
    Dim wbData As Workbook: Set wbData = openExcelFile()
    wbData.Activate
    
    Call RunQueries(accessPath, ExcelToAccessSQL(tblName))
    
    wbData.Close

    MsgBox ("Upload Complete")

blank:
    Call optimize(False)
End Sub

Sub getAccessPath()
    
    On Error GoTo reason
    
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd.Filters
        .Clear
        .Add "Access files", "*.accdb"
    End With
    
    fd.Show
    ThisWorkbook.Sheets(1).Range("D3").Value = fd.SelectedItems(1)

    Exit Sub
    
reason:
    MsgBox "No file selected"
    End
End Sub

'==============================================================================
'MAIN FUNCTIONS
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
    ReDim sqlstring(part + 4) As String
    
    'Drop / Create Table Scripts
    sqlstring(0) = "DROP TABLE [" & tblName & "]"
    sqlstring(1) = "CREATE TABLE [" & tblName & "] ("
    
    For i = 0 To intlastCol - 1
        With srcWS
            sqlstring(1) = sqlstring(1) & "[" & .Range("A1").Offset(0, i).Value & "] MEMO, "
        End With
    Next i
    
    sqlstring(1) = Left(sqlstring(1), Len(sqlstring(1)) - 2) & ")"
    
    '!!DEBUG: Copy Create Table script in Cell A1
    'ThisWorkbook.Sheets(1).Range("A1").Value = sqlstring(1)
    
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
        srcWS.Range("A1:" & lastColumn & 60001).NumberFormat = "@"
        sqlstring(2) = sqlinsert & sqlselect & sqlfrom & sqlsheet
        
        Dim tempWS As Worksheet
        For i = 3 To part + 1
            Set tempWS = srcWB.Sheets.Add
            srcWS.Range("A1:" & lastColumn & "1").Copy tempWS.Range("A1")
            srcWS.Range("A" & (60001 * (i - 2)) + 1 & ":" & lastColumn & (60001 * (i - 1))).Copy    
            tempWS.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
            Selection.NumberFormat = "@"
            Application.CutCopyMode = False
            lastRow = tempWS.Cells(Rows.Count, 1).End(xlUp).Row
            
            sqlsheet = "[" & tempWS.Name & "$A1:" & lastColumn & lastRow & "]"
            sqlstring(i) = sqlinsert & sqlselect & sqlfrom & sqlsheet
        Next i
        
    End If
    
    'Add [UploadDate] and [ReportDate] to the data
    sqlstring(part + 2) = "ALTER TABLE " & tblName & " ADD COLUMN [UploadDate] DATE"
    sqlstring(part + 3) = "ALTER TABLE " & tblName & " ADD COLUMN [ReportDate] DATE"
    sqlstring(part + 4) = "UPDATE " & tblName & " Set [UploadDate] = #" & Format(Now, "mm/dd/yyyy") & _
        "#,[ReportDate] = #" & Format(reportDate, "mm/dd/yyyy") & "#"
    
    ExcelToAccessSQL = sqlstring
End Function

'==============================================================================
'UTILITY FUNCTIONS
'==============================================================================

Function openExcelFile() As Workbook

    MsgBox ("Please select excel file to load")
    On Error GoTo blank
    
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd.Filters
        .Clear
        .Add "Excel files", "*.xlsx;*.xls;*.xlsm;*.xlsb"
    End With
    
    fd.Show
    Set openExcelFile = Workbooks.Open(fd.SelectedItems(1), False)
    
    Exit Function
    
blank:
    MsgBox "No file selected"
    End
End Function

Sub RunQueries(accessFilePath As String, sqlstring() As String)

    On Error GoTo blank
    Dim scn As String: scn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & accessFilePath
    Set cn = CreateObject("ADODB.Connection")
    cn.Open scn
    
    For i = 1 To UBound(sqlstring)
        cn.Execute sqlstring(i)
    Next i
    cn.Close
    
    Exit Sub

blank:
    If Err.Number <> -2147217900 Then
        MsgBox (Err.Description)
        End
    End If
    
    For i = 0 To UBound(sqlstring)
        cn.Execute sqlstring(i)
    Next i
    cn.Close

End Sub

Function getlastRow(columnRef As String) As Long

    With ActiveWorkbook.ActiveSheet
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

Sub DeleteAllSheets()
    Application.DisplayAlerts = False
    
    ThisWorkbook.Sheets(1).Activate
    
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name = ThisWorkbook.ActiveSheet.Name Then
            sht.Cells.Delete
        Else
            sht.Delete
        End If
    Next sht
        
    ThisWorkbook.Sheets(1).Name = "Sheet1"
    ThisWorkbook.VBProject.VBComponents(Sheets(1).CodeName).Name = "Sheet1"
    
    Application.DisplayAlerts = True
End Sub

'==============================================================================
'LICENSE
'==============================================================================

'MIT License
'
'Copyright (c) 2019 Ampert - glenn.marvin.l.lim@accenture.com
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

