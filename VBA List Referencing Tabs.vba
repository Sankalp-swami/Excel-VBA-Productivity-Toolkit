Sub FindReferencingTabs()
    Dim ws As Worksheet
    Dim tabName As String
    Dim resultSheet As Worksheet
    Dim outputRow As Long
    Dim cell As Range
    Dim foundTabs As Object
    Dim tabAlreadyAdded As Boolean
    
    ' Define the target tab name
    tabName = "Project List" ' Replace with the exact name of the tab you want to check
    
    ' Initialize a dictionary to store unique tab names
    Set foundTabs = CreateObject("Scripting.Dictionary")
    
    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> tabName Then
            ' Check each cell in UsedRange for formula references
            For Each cell In ws.UsedRange
                If Not IsError(cell) Then
                    If InStr(1, cell.Formula, "'" & tabName & "'!", vbTextCompare) > 0 Or _
                       InStr(1, cell.Formula, tabName & "!", vbTextCompare) > 0 Then
                        ' Add the referencing tab name to the dictionary
                        If Not foundTabs.exists(ws.Name) Then
                            foundTabs.Add ws.Name, True
                        End If
                        Exit For ' Stop scanning this sheet once we find a reference
                    End If
                End If
            Next cell
        End If
    Next ws
    
    ' Write results to a new sheet
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Worksheets("Referencing Tabs")
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Worksheets.Add
        resultSheet.Name = "Referencing Tabs"
    End If
    On Error GoTo 0
    
    ' Clear previous results
    resultSheet.Cells.Clear
    resultSheet.Cells(1, 1).Value = "Tabs Referencing " & tabName
    outputRow = 2
    
    ' Write each referencing tab name to the result sheet
    Dim key As Variant
    For Each key In foundTabs.keys
        resultSheet.Cells(outputRow, 1).Value = key
        outputRow = outputRow + 1
    Next key
    
    ' Notify user
    If foundTabs.Count = 0 Then
        MsgBox "No tabs reference " & tabName & "."
    Else
        MsgBox "Tabs referencing " & tabName & " are listed in the 'Referencing Tabs' sheet."
    End If
End Sub
