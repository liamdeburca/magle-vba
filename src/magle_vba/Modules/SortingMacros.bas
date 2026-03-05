'===============================================================================
' Module: SortingMacros
'===============================================================================
' Description:
'   Macros for sorting data rows in the Data worksheet.
'   Provides ascending and descending sort operations that rearrange all
'   data rows based on values in a selected row.
'===============================================================================

'===============================================================================
' [FUNCTION] GetActiveDataRow
'===============================================================================
' Description:
'   Retrieves the DataRowCls instance corresponding to the row of the
'   specified cell by looking up its key value.
'
' Parameters:
'   cell : Range
'       Any cell in the target row
'
' Returns:
'   DataRowCls - The data row matching the cell's row key
'===============================================================================
Private Function GetActiveDataRow( _
    cell As Range _
) As DataRowCls
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")

    Dim key As String
    key = ws.Range(Specs.KeyColumn & cell.Row).value

    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()

    Dim dataRow As DataRowCls
    Set dataRow = ParsedData.GetRowFromKey(key)

    Set GetActiveDataRow = dataRow
End Function

'===============================================================================
' [SUB] SortAscending
'===============================================================================
' Description:
'   Sorts all data rows based on values in the row containing the specified
'   cell, in ascending order. Skips if already sorted ascending.
'
' Parameters:
'   cell : Range
'       Any cell in the row to sort by
'===============================================================================
Public Sub SortAscending( _
    cell As Range _
)
    Dim dataRow As DataRowCls
    Set dataRow = GetActiveDataRow(cell)

    If dataRow.IsSorted(ascending:=True) Then
        Exit Sub
    End If

    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()

    Call ParsedData.SortAlongRow(dataRow.key, ascending:=True)
End Sub

'===============================================================================
' [SUB] SortDescending
'===============================================================================
' Description:
'   Sorts all data rows based on values in the row containing the specified
'   cell, in descending order. Skips if already sorted descending.
'
' Parameters:
'   cell : Range
'       Any cell in the row to sort by
'===============================================================================
Public Sub SortDescending( _
    cell As Range _
)
    Dim dataRow As DataRowCls
    Set dataRow = GetActiveDataRow(cell)

    If dataRow.IsSorted(ascending:=False) Then
        Exit Sub
    End If

    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()

    Call ParsedData.SortAlongRow(dataRow.key, ascending:=False)
End Sub
