Private Function GetActiveDataRow(cell As Range) As DataRowCls
    ' Gets the corresponding data row to the specified cell
    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()

    Set GetActiveDataRow = ParsedData.GetRowFromIndex(cell.Row)
End Function

Public Sub ApplyMacrosToDataRow(dataRow As DataRowCls)
    ' Applies a macro dropdown menu to the macro-cell in each datarow
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()

    Dim rng As Range
    Set rng = ws.Range(Specs.MacroColumn & dataRow.rowIdx)

    ' Clear any existing validation first
    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0

    ' Add the new validation
    With rng.Validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="Describe,Sort (ASC),Sort (DESC),Scatter"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Select a macro to run"
        .ErrorTitle = "Invalid macro!"
    End With
End Sub

Public Sub RunMacroFromDropdown(cell As Range)
    ' Runs a macro specified by the cell, and subsequently clears the cell value
    Dim cellValue As Variant
    cellValue = cell.value
    
    If IsEmpty(cellValue) Or CStr(cellValue) = "" Then
        Exit Sub
    End If

    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()
    
    Dim dataRow As DataRowCls
    Set dataRow = GetActiveDataRow(cell)

    Select Case CStr(cell)
        Case "Describe"
            Call dataRow.Describe
        
        Case "Sort (ASC)"
            If Not dataRow.IsSorted(ascending:=True) Then
                Call ParsedData.SortAlongRow(dataRow.key, ascending:=True)
            End If
            
        Case "Sort (DESC)"
            If Not dataRow.IsSorted(ascending:=False) Then
                Call ParsedData.SortAlongRow(dataRow.key, ascending:=False)
            End If
            
        Case "Scatter"
            Call PlottingMacros.BasicScatterPlot_(cell)
        Case Else
            MsgBox "Macro " & macro & "not implemented yet!"
    End Select
    
    cell.value = ""
End Sub
