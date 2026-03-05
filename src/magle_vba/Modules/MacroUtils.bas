'===============================================================================
' Module: MacroUtils
'===============================================================================
' Description:
'   Utilities for managing macro dropdown menus in the Data sheet.
'   Provides functionality to add dropdown validation cells and execute
'   user-selected macros such as Describe, Sort, and Scatter plot.
'===============================================================================

'===============================================================================
' [FUNCTION] GetActiveDataRow
'===============================================================================
' Description:
'   Retrieves the DataRowCls instance corresponding to a given cell's row.
'
' Parameters:
'   cell : Range
'       Any cell in the target row
'
' Returns:
'   DataRowCls - The data row at the cell's row index
'===============================================================================
Private Function GetActiveDataRow( _
    cell As Range _
) As DataRowCls
    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()

    Set GetActiveDataRow = ParsedData.GetRowFromIndex(cell.Row)
End Function

'===============================================================================
' [SUB] ApplyMacrosToDataRow
'===============================================================================
' Description:
'   Adds a dropdown validation menu to the macro column cell for a data row.
'   The dropdown contains options: Describe, Sort (ASC), Sort (DESC), Scatter.
'
' Parameters:
'   dataRow : DataRowCls
'       The data row to add the macro dropdown to
'===============================================================================
Public Sub ApplyMacrosToDataRow( _
    dataRow As DataRowCls _
)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()

    Dim rng As Range
    Set rng = ws.Range(Specs.MacroColumn & dataRow.rowIdx)

    On Error Resume Next
    rng.Validation.Delete
    On Error GoTo 0

    With rng.Validation
        .Add _
            Type:=xlValidateList, _
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
'   Executes the macro specified by a dropdown cell value, then clears the
'   cell. Supports Describe, Sort (ASC), Sort (DESC), and Scatter operations.
'
' Parameters:
'   cell : Range
'       The dropdown cell containing the selected macro name
'===============================================================================
Public Sub RunMacroFromDropdown( _
    cell As Range _
)
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
            MsgBox "Macro " & CStr(cellValue) & " not implemented yet!"
    End Select
    
    cell.value = ""
End Sub
