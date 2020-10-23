'TODO:
'Last Column Finder
'Last Row Finder

'==============================================================================
'EXCEL TO EXCEL COPY PASTER
'==============================================================================
'Take any excel file and copy all values with just a few clicks!

'==============================================================================
'REQUIREMENTS AND LIMITATIONS
'==============================================================================
'1. Target excel file must only contain 1 sheet
'2. Data will only be copied to the file that contains this module
'3. Merged cells will only consider the master cell value and have the rest as blank

'==============================================================================
'INSTALLATION GUIDE
'==============================================================================
'Run Install()

'==============================================================================
'MAIN
'==============================================================================

Sub CopyPaste()
    Dim rawdata as Workbook
    Set rawdata = openExcelFile

    ThisWorkbook.Sheets(1).Range("A:AA").Value = rawdata.Sheets(1).Range("A:AA").Value

End Sub

'==============================================================================
'UTILITIES
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