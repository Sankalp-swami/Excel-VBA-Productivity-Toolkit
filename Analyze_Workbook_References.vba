Sub AnalyzeWorkbookReferences()
    Dim ws As Worksheet, sourceSheet As Worksheet, resultSheet As Worksheet
    Dim cell As Range
    Dim formula As String
    Dim referencingTabs As Object, referencedTabs As Object
    Dim tabReferences As Object
    Dim key As Variant, match As Variant
    Dim outputRow As Long

    ' Initialize dictionary to store all references
    Set tabReferences = CreateObject("Scripting.Dictionary")

    ' Analyze each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Initialize dictionaries for each tab
        Set referencingTabs = CreateObject("Scripting.Dictionary")
        Set referencedTabs = CreateObject("Scripting.Dictionary")

        ' Scan all formulas in the current sheet
        For Each cell In ws.UsedRange
            If Not IsError(cell) And cell.HasFormula Then
                formula = cell.Formula

                ' Get all tab references in the formula
                Dim matches As Variant
                matches = GetSheetReferences(formula)

                ' Add current tab references
                For Each match In matches
                    If WorksheetExists(CStr(match)) Then
                        ' Add to referenced tabs if not already present
                        If Not referencedTabs.exists(CStr(match)) Then
                            referencedTabs.Add CStr(match), True
                        End If
                    End If
                Next match
            End If
        Next cell

        ' Add referencing tabs for all other sheets
        For Each sourceSheet In ThisWorkbook.Worksheets
            If sourceSheet.Name <> ws.Name Then
                For Each cell In sourceSheet.UsedRange
                    If Not IsError(cell) And cell.HasFormula Then
                        formula = cell.Formula
                        If InStr(1, formula, "'" & ws.Name & "'!", vbTextCompare) > 0 Or _
                           InStr(1, formula, ws.Name & "!", vbTextCompare) > 0 Then
                            If Not referencingTabs.exists(sourceSheet.Name) Then
                                referencingTabs.Add sourceSheet.Name, True
                            End If
                        End If
                    End If
                Next cell
            End If
        Next sourceSheet

        ' Store results for the current tab
        tabReferences.Add ws.Name, Array(referencingTabs.keys, referencedTabs.keys)
    Next ws

    ' Write results to a new sheet
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Worksheets("Workbook References Summary")
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Worksheets.Add
        resultSheet.Name = "Workbook References Summary"
    End If
    On Error GoTo 0

    ' Clear previous results
    resultSheet.Cells.Clear
    resultSheet.Cells(1, 1).Value = "Tab Name"
    resultSheet.Cells(1, 2).Value = "Tabs Referencing This Tab"
    resultSheet.Cells(1, 3).Value = "Tabs Referenced by This Tab"
    outputRow = 2

    ' Output results
    For Each key In tabReferences.keys
        resultSheet.Cells(outputRow, 1).Value = key
        resultSheet.Cells(outputRow, 2).Value = Join(tabReferences(key)(0), ", ")
        resultSheet.Cells(outputRow, 3).Value = Join(tabReferences(key)(1), ", ")
        outputRow = outputRow + 1
    Next key

    ' Notify user
    MsgBox "Workbook analysis complete. Results are available in 'Workbook References Summary' sheet."
End Sub

' Helper function to extract sheet references from a formula
Function GetSheetReferences(formula As String) As Variant
    Dim regex As Object
    Dim matches As Object
    Dim match As Variant
    Dim results() As String
    Dim i As Long

    ' Regular expression to capture sheet references (e.g., SheetName!)
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "'([^']+)'!|([A-Za-z0-9_]+)!"
    regex.Global = True

    ' Initialize results array
    ReDim results(0)

    ' Extract matches
    If regex.Test(formula) Then
        Set matches = regex.Execute(formula)
        For Each match In matches
            If match.SubMatches(0) <> "" Then
                ReDim Preserve results(UBound(results) + 1)
                results(UBound(results)) = match.SubMatches(0)
            ElseIf match.SubMatches(1) <> "" Then
                ReDim Preserve results(UBound(results) + 1)
                results(UBound(results)) = match.SubMatches(1)
            End If
        Next match
    End If

    ' Return unique sheet names
    GetSheetReferences = results
End Function

' Helper function to check if a worksheet exists
Function WorksheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function
