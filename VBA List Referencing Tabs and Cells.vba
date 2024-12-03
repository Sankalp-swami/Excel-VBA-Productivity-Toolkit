Sub FindTabReferences()
    Dim ws As Worksheet
    Dim cell As Range
    Dim tabName As String
    Dim resultSheet As Worksheet
    Dim outputRow As Long
    Dim formula As String
    
    tabName = "Forecast Changes" 'Replace with the exact name of the tab you want to check
    
    ' Create a new sheet for results
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Worksheets("Tab References")
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Worksheets.Add
        resultSheet.Name = "Tab References"
    End If
    On Error GoTo 0
    
    ' Clear the result sheet
    resultSheet.Cells.Clear
    resultSheet.Cells(1, 1).Value = "Referencing Tab"
    resultSheet.Cells(1, 2).Value = "Cell Address"
    resultSheet.Cells(1, 3).Value = "Formula"
    outputRow = 2
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> tabName Then
            ' Loop through used cells in each worksheet
            For Each cell In ws.UsedRange
                If Not IsError(cell) Then
                    formula = cell.Formula
                    ' Check for explicit references
                    If InStr(1, formula, "'" & tabName & "'!", vbTextCompare) > 0 Or _
                       InStr(1, formula, tabName & "!", vbTextCompare) > 0 Then
                        resultSheet.Cells(outputRow, 1).Value = ws.Name
                        resultSheet.Cells(outputRow, 2).Value = cell.Address
                        resultSheet.Cells(outputRow, 3).Value = formula
                        outputRow = outputRow + 1
                    End If
                    ' Check for structured references
                    If InStr(1, formula, "[" & tabName & "]", vbTextCompare) > 0 Then
                        resultSheet.Cells(outputRow, 1).Value = ws.Name
                        resultSheet.Cells(outputRow, 2).Value = cell.Address
                        resultSheet.Cells(outputRow, 3).Value = formula
                        outputRow = outputRow + 1
                    End If
                End If
            Next cell
        End If
    Next ws
    
    ' Notify user
    If outputRow = 2 Then
        MsgBox "No references to " & tabName & " found."
    Else
        MsgBox "References to " & tabName & " have been written to the 'Tab References' sheet."
    End If
End Sub
