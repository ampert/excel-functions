References:
OLE Automation
Microsoft ActiveX Data Objects 2.8 Library
Microsoft Access 16.0 Object Library

Sub SEAOpenItems()
    Application.ScreenUpdating = False
    Dim fd As FileDialog
    Dim FileChosen As Integer
    Dim FileName As String
    Dim refLastRow As Long
    Dim fileLastRow As Long
    Dim i As Long, x As Long, c As Long, lastRow As Long, lastCol As Long
    Dim cn As Variant, scn As String
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    'fd.InitialFileName = ThisWorkbook.Path
    fd.InitialView = msoFileDialogViewList
    fd.AllowMultiSelect = False 'True
    FileChosen = fd.Show
    
    Dim ws As Worksheet
    Dim wsTemp As Worksheet
    Dim fileSh As Worksheet
    Dim wbReference As Workbook
    Dim dbWB As String, dsh As String, ssql As String
    
    Set fileSh = ThisWorkbook.Sheets("RUN INFO")
    fileLastRow = fileSh.Cells(Rows.Count, 4).End(xlUp).Row + 1
    
    Set cn = CreateObject("ADODB.Connection")
    dbpath = ThisWorkbook.Path & "\Lenovo DB_v3.accdb"
    scn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbpath
    
    If FileChosen = -1 Then
     startTime = Now
        cn.Open scn
        ssql = "DELETE * FROM SEAOpenItems"
        cn.Execute ssql
        For i = 1 To fd.SelectedItems.Count
            
            Set wbReference = Workbooks.Open(fd.SelectedItems(i), False)
            
            fileSh.Range("A" & fileLastRow) = fileLastRow
            fileSh.Range("B" & fileLastRow) = Date
            fileSh.Range("C" & fileLastRow) = DataSource
            fileSh.Range("D" & fileLastRow) = wbReference.FullName
            fileSh.Range("E" & fileLastRow) = ReportingDate
            
            fileLastRow = fileSh.Cells(Rows.Count, 4).End(xlUp).Row + 1
            
            Set ws = wbReference.Sheets("Sheet1") '("Name of Sheet")
            ws.Columns.EntireColumn.Hidden = False
            If ws.AutoFilterMode Then ws.AutoFilterMode = False
            refLastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            lastRow = refLastRow
            lastCol = ws.Cells(3, Columns.Count).End(xlToLeft).Column
            
            
            If refLastRow > 1 And refLastRow <= 60000 Then
            
                dbWB = wbReference.FullName
                dsh = "[" & ws.Name & "$A1:BZ" & lastRow & "]"
                ssql = "INSERT INTO SEAOpenItems "
                ssql = ssql & _
                            "SELECT [BG], [GEO], [Region], Country, [Company Code], [Customer ID], " & _
                            "[Customer Name], [Doc_Currency], [Document Type], [Document_date], [Document_number], " & _
                            "[Baseline_date], [Payment_Terms], [Due_date], [Posting_date], [AR Balance], " & _
                            "[Exchange_Rate], [AR_local_amount] AS [AR Local amount], [Clearing_date], [Owner], " & _
                            IIf(MonthEnd, 1, 0) & " AS [Month End], '" & ReportingDate & "' AS [Reporting Date] " & _
                            "FROM [Excel 8.0;HDR=YES;DATABASE=" & dbWB & "]." & dsh & " WHERE [Country] IN ('SG', 'MY', 'VN', 'PH', 'KH', 'MM')"
                cn.Execute ssql
            
            ElseIf refLastRow > 60000 Then
                
                Dim partial As Integer
                partial = roundingOff(refLastRow)
                                
                dbWB = wbReference.FullName
                dsh = "[" & ws.Name & "$A1:BZ60001]"
                ssql = "INSERT INTO SEAOpenItems "
                ssql = ssql & _
                            "SELECT [BG], [GEO], [Region], Country, [Company Code], [Customer ID], " & _
                            "[Customer Name], [Doc_Currency], [Document Type], [Document_date], [Document_number], " & _
                            "[Baseline_date], [Payment_Terms], [Due_date], [Posting_date], [AR Balance], " & _
                            "[Exchange_Rate], [AR_local_amount] AS [AR Local amount], [Clearing_date], [Owner], " & _
                            IIf(MonthEnd, 1, 0) & " AS [Month End], '" & ReportingDate & "' AS [Reporting Date] " & _
                            "FROM [Excel 8.0;HDR=YES;DATABASE=" & dbWB & "]." & dsh & " WHERE [Country] IN ('SG', 'MY', 'VN', 'PH', 'KH', 'MM')"
                cn.Execute ssql
                
                    For x = 2 To partial
                        Set wsTemp = wbReference.Sheets.Add
                        ws.Range("A1:BZ1").Copy wsTemp.Range("A1")
                        ws.Range("A" & (60001 * (x - 1)) + 1 & ":BZ" & (60001 * x)).Copy
                        
                        wsTemp.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
                        Application.CutCopyMode = False
                        refLastRow = wsTemp.Cells(Rows.Count, 1).End(xlUp).Row
                        lastRow = refLastRow
                                        
                        dbWB = wbReference.FullName
                        dsh = "[" & wsTemp.Name & "$A1:BZ" & lastRow & "]"
                        ssql = "INSERT INTO SEAOpenItems "
                        ssql = ssql & _
                            "SELECT [BG], [GEO], [Region], Country, [Company Code], [Customer ID], " & _
                            "[Customer Name], [Doc_Currency], [Document Type], [Document_date], [Document_number], " & _
                            "[Baseline_date], [Payment_Terms], [Due_date], [Posting_date], [AR Balance], " & _
                            "[Exchange_Rate], [AR_local_amount] AS [AR Local amount], [Clearing_date], [Owner], " & _
                            IIf(MonthEnd, 1, 0) & " AS [Month End], '" & ReportingDate & "' AS [Reporting Date] " & _
                            "FROM [Excel 8.0;HDR=YES;DATABASE=" & dbWB & "]." & dsh & " WHERE [Country] IN ('SG', 'MY', 'VN', 'PH', 'KH', 'MM')"
                        cn.Execute ssql
                    Next x
            End If
            
            wbReference.Close False
            Set wbReference = Nothing
            Set ws = Nothing
        Next i
        cn.Close
        CompletedTask = True
    Else
        CompletedTask = False
        
    End If
Application.ScreenUpdating = True

    Exit Sub
ErrHandler:
    CompletedTask = False
    If Err.Number = -2147467259 Then
        MsgBox "Process enterupted! One or more required column is missing."
    Else
        MsgBox "Process enterupted due to the following error: " & Err.Number & " - " & Err.Description
    End If
    'wbReference.Close False
    Application.ScreenUpdating = True
    
End Sub

Function roundingOff(lastrowPartial As Long) As Long
    Dim remainder As Long
    Dim firstBatch As Integer
    remainder = lastrowPartial Mod 60000
    
    firstBatch = Int(lastrowPartial / 60000)
    
    If remainder > 0 Then
    firstBatch = firstBatch + 1
    End If
    roundingOff = firstBatch
    'MsgBox firstBatch
End Function
