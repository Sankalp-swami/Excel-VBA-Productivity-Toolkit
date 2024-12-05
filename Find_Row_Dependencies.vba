Sub FindRowDependencies()
    Dim ws As Worksheet
    Dim targetTab As Worksheet
    Dim targetCol As String
    Dim searchTerm As String
    Dim foundRefs As String
    Dim rowCell As Range
    Dim sheetCell As Range
    Dim rowRef As String
    Dim outputSheet As Worksheet
    Dim outputRow As Long
    
    ' Define parameters
    targetCol = "D" ' Column to check
    searchTerm = "Rollup" ' Term to search for
    Set targetTab = ThisWorkbook.Worksheets("Project List") ' Tab to search in

    ' Initialize output
    foundRefs = ""
    On Error Resume Next
    Set outputSheet = ThisWorkbook.Worksheets("Row Dependencies")
    If outputSheet Is Nothing Then
        Set outputSheet = ThisWorkbook.Worksheets.Add
        outputSheet.Name = "Row Dependencies"
    End If
    On Error GoTo 0
    outputSheet.Cells.Clear
    outputSheet.Cells(1, 1).Value = "Referencing Sheet"
    outputSheet.Cells(1, 2).Value = "Cell Address"
    outputSheet.Cells(1, 3).Value = "Row Reference"
    outputRow = 2

    ' Iterate through the target column
    For Each rowCell In targetTab.Columns(targetCol).Cells
        If Not IsError(rowCell.Value) And Not IsEmpty(rowCell.Value) Then
            If rowCell.Value = searchTerm Then
                ' Create row reference string
                rowRef = "'" & targetTab.Name & "'!" & rowCell.EntireRow.Address

                ' Check dependencies in other sheets
                For Each ws In ThisWorkbook.Worksheets
                    If ws.Name <> targetTab.Name Then
                        For Each sheetCell In ws.UsedRange
                            If Not IsError(sheetCell) And Not IsEmpty(sheetCell) Then
                                If VarType(sheetCell.Formula) = vbString Then
                                    If InStr(1, sheetCell.Formula, rowRef, vbTextCompare) > 0 Then
                                        ' Log results to output sheet
                                        outputSheet.Cells(outputRow, 1).Value = ws.Name
                                        outputSheet.Cells(outputRow, 2).Value = sheetCell.Address
                                        outputSheet.Cells(outputRow, 3).Value = rowRef
                                        outputRow = outputRow + 1
                                    End If
                                End If
                            End If
                        Next sheetCell
                    End If
                Next ws
            End If
        End If
    Next rowCell

    ' Display message
    If outputRow = 2 Then
        MsgBox "No dependencies found for rows with '" & searchTerm & "' in column " & targetCol & ".", vbInformation
    Else
        MsgBox "Dependencies found! Check the 'Row Dependencies' tab for details.", vbInformation
    End If
End Sub
