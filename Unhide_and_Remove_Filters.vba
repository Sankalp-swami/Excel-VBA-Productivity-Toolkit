Sub UnhideAndRemoveFilters()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long

    ' Turn off screen updating and recalculations for speed
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each ws In ThisWorkbook.Worksheets
        ' Unhide all rows and columns
        ws.Rows.Hidden = False
        ws.Columns.Hidden = False
        
        ' Expand all grouped rows and columns
        ws.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8

        ' Remove filters
        If ws.AutoFilterMode Then
            ws.AutoFilterMode = False
        End If
    Next ws

    ' Restore screen updating and recalculations
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "All filters removed, and all rows/columns are unhidden and expanded!"
End Sub
