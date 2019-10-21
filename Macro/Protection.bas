Sub UnProtectAllSheets(pwd as string)

    Dim xWs As Worksheet
    For Each xWs In Application.ActiveWorkbook.Worksheets
        xWs.Unprotect pwd
    Next
    
End Sub

Sub Locker(islock As Boolean, sheetName As String, pwd As String)

    If islock Then
        ThisWorkbook.Sheets(sheetName).Protect pwd, DrawingObjects:=True, Contents:=True, Scenarios:=False, AllowFiltering:=True
    Else
        ThisWorkbook.Sheets(sheetName).Unprotect pwd
    End If

End Sub

Sub LockBook(pwd as string)

    ThisWorkbook.Protect pwd, Structure:=True, Windows:=False

End Sub