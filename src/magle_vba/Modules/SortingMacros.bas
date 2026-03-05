' File containing macros for sorting data in Exce

Private Function GetActiveDataRow(cell As Range) As DataRowCls
    ' Gets the corresponding data row to the active cell
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

Public Sub SortAscending(cell As Range)
    Dim dataRow As DataRowCls
    Set dataRow = GetActiveDataRow(cell)

    If dataRow.IsSorted(ascending:=True) Then
        Exit Sub
    End If

    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()

    Call ParsedData.SortAlongRow(dataRow.key, ascending:=True)
End Sub

Public Sub SortDescending(cell As Range)
    Dim dataRow As DataRowCls
    Set dataRow = GetActiveDataRow(cell)

    If dataRow.IsSorted(ascending:=False) Then
        Exit Sub
    End If

    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()

    Call ParsedData.SortAlongRow(dataRow.key, ascending:=False)
End Sub
