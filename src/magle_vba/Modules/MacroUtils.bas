'===============================================================================
' [FUNCTION] GetActiveDataRow
'===============================================================================
' Description:
'   Returns the DataRowCls instance that corresponds to the worksheet row
'   of the specified cell. Looks up the row by its index in the current
'   ParsedDataCls instance.
'
' Parameters:
'   cell : Range
'       Any cell in the Data worksheet; its row number is used as the key
'
' Returns:
'   DataRowCls
'       The DataRowCls instance for that worksheet row
'===============================================================================
Private Function GetActiveDataRow(cell As Range) As DataRowCls
    ' Gets the corresponding data row to the specified cell
    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()

    Set GetActiveDataRow = ParsedData.GetRowFromIndex(cell.Row)
End Function

'===============================================================================
' [SUB] ApplyMacrosToDataRow
'===============================================================================
' Description:
'   Adds an in-cell dropdown validation list to the macro column cell of
'   the given DataRowCls instance. The dropdown offers these actions:
'   Describe, Sort (ASC), Sort (DESC), Scatter. When the user selects an
'   option, SheetData.Worksheet_Change fires and delegates to
'   RunMacroFromDropdown.
'
' Parameters:
'   dataRow : DataRowCls
'       The row whose macro cell should receive the dropdown
'
' Notes:
'   - Any existing validation on the cell is removed before adding a new
'     one to avoid duplicates
'===============================================================================
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

'===============================================================================
' [SUB] RunMacroFromDropdown
'===============================================================================
' Description:
'   Executes the action selected in a macro dropdown cell and then clears
'   the cell so it returns to its default empty state. Supported actions:
'     - Describe    : Shows a statistical summary for the row
'     - Sort (ASC)  : Sorts all rows ascending by this row's values
'     - Sort (DESC) : Sorts all rows descending by this row's values
'     - Scatter     : Creates a basic scatter plot for this row
'
' Parameters:
'   cell : Range
'       The macro column cell that contains the selected action string
'
' Notes:
'   - Called by SheetData.Worksheet_Change when the macro column changes
'   - Does nothing if the cell is empty
'===============================================================================
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
            Call BasicScatterPlot_(cell)
        Case Else
            MsgBox "Macro " & macro & "not implemented yet!"
    End Select
    
    cell.value = ""
End Sub
