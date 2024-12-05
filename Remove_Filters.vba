Sub RemoveAllFilters()
    Application.ScreenUpdating = False ' Turn off screen updates
    Application.Calculation = xlCalculationManual ' Turn off recalculations
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
    Next ws
    
    Application.Calculation = xlCalculationAutomatic ' Restore recalculations
    Application.ScreenUpdating = True ' Turn screen updates back on
    MsgBox "All filters removed!"
End Sub