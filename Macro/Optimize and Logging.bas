Sub optilog(processName As String, start As Boolean)
    lastrow = ThisWorkbook.Sheets("LOGS").Range("A" & Rows.Count).End(xlUp).Row + 1
    
    If start Then
        Application.ScreenUpdating = False
        Application.DisplayStatusBar = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        ThisWorkbook.Sheets("LOGS").Range("A" & lastrow).Value = processName
        ThisWorkbook.Sheets("LOGS").Range("B" & lastrow).Value = Now
    Else
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        lastrow = lastrow - 1
        ThisWorkbook.Sheets("LOGS").Range("C" & lastrow).Value = Now
    End If

End Sub

Sub optimize(start As Boolean)

    If start Then
        Application.ScreenUpdating = False
        Application.DisplayStatusBar = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
    Else
        Application.ScreenUpdating = True
        Application.DisplayStatusBar = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
    End If

End Sub

Sub logger(processName As String, start As Boolean)
    lastrow = ThisWorkbook.Sheets("LOGS").Range("A" & Rows.Count).End(xlUp).Row + 1
    
    If start Then
        ThisWorkbook.Sheets("LOGS").Range("A" & lastrow).Value = processName
        ThisWorkbook.Sheets("LOGS").Range("B" & lastrow).Value = Now
    Else
        ThisWorkbook.Sheets("LOGS").Range("C" & lastrow).Value = Now
    End If

End Sub