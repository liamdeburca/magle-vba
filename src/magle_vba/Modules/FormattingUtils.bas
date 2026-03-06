' Utilities For formatting values in the Data sheet

'===============================================================================
' [FUNCTION] IsBelowMin
'===============================================================================
' Description:
'   Returns True if the numeric value of the given cell is less than the
'   minimum bound defined for that measurement row. Called by Excel
'   conditional formatting formulas to highlight out-of-range values in
'   blue.
'
' Parameters:
'   cell : Range
'       The cell whose value is being evaluated
'
' Returns:
'   Boolean
'       True if the cell value is below the row minimum, False otherwise
'
' Notes:
'   - Returns False for empty cells or cells without a defined minimum
'   - Relies on ParsedDataCls to look up the DataRowCls for the cell's row
'===============================================================================
Public Function IsBelowMin(cell As Range) As Boolean
    If IsEmpty(cell.value) Or cell.value = "" Then
        IsBelowMin = False
        Exit Function
    End If

    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()

    Dim dataRow As DataRowCls
    Set dataRow = ParsedData.GetRowFromIndex(cell.Row)

    If IsError(dataRow.min) Then
        IsBelowMin = False
        Exit Function
    End If

    IsBelowMin = (CDbl(cell.value) < CDbl(dataRow.min))
End Function

'===============================================================================
' [FUNCTION] IsAboveMax
'===============================================================================
' Description:
'   Returns True if the numeric value of the given cell exceeds the maximum
'   bound defined for that measurement row. Called by Excel conditional
'   formatting formulas to highlight out-of-range values in red.
'
' Parameters:
'   cell : Range
'       The cell whose value is being evaluated
'
' Returns:
'   Boolean
'       True if the cell value is above the row maximum, False otherwise
'
' Notes:
'   - Returns False for empty cells or cells without a defined maximum
'   - Relies on ParsedDataCls to look up the DataRowCls for the cell's row
'===============================================================================
Public Function IsAboveMax(cell As Range) As Boolean
    If IsEmpty(cell.value) Or cell.value = "" Then
        IsAboveMax = False
        Exit Function
    End If

    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()

    Dim dataRow As DataRowCls
    Set dataRow = ParsedData.GetRowFromIndex(cell.Row)

    If IsError(dataRow.max) Then
        IsAboveMax = False
        Exit Function
    End If

    IsAboveMax = (CDbl(cell.value) > CDbl(dataRow.max))
End Function

'===============================================================================
' [SUB] ApplyConditionalFormattingToDataRow
'===============================================================================
' Description:
'   Writes two conditional formatting rules to the full data range of the
'   given DataRowCls row in the Data worksheet:
'     - Values below the row minimum are shown in bold blue text
'     - Values above the row maximum are shown in bold red text
'   The rules use formula-based conditions that call IsBelowMin and
'   IsAboveMax respectively, so formatting updates dynamically as data
'   changes.
'
' Parameters:
'   dataRow : DataRowCls
'       The row instance whose worksheet range should receive the rules
'
' Notes:
'   - Any existing formatting rules on the range are deleted first
'   - Called by DataRowCls.ApplyConditionalFormatting during Init
'===============================================================================
Public Sub ApplyConditionalFormattingToDataRow(dataRow As DataRowCls)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()

    Dim dataStartCol As Long
    dataStartCol = ws.Range(Specs.DataStartColumn & "1").Column

    Dim firstCell As Range
    Set firstCell = ws.Cells(dataRow.rowIdx, dataStartCol)
    Dim firstCellRef As String
    firstCellRef = firstCell.Address(False, False)  ' Relative reference like A5

    Dim dataRange As Range
    Set dataRange = ws.Range(ws.Cells(dataRow.rowIdx, dataStartCol), _
        ws.Cells(dataRow.rowIdx, dataStartCol + Specs.NumColumns - 1))

    dataRange.FormatConditions.Delete

    ' Add "Below Min" rule - formula evaluates each cell individually
    With dataRange.FormatConditions.Add(xlExpression, , "=IsBelowMin(" & firstCellRef & ")")
        With .Font
            .Color = RGB(0, 0, 255)  ' Blue
            .Bold = True
        End With
    End With

    ' Add "Above Max" rule - formula evaluates each cell individually
    With dataRange.FormatConditions.Add(xlExpression, , "=IsAboveMax(" & firstCellRef & ")")
        With .Font
            .Color = RGB(255, 0, 0)  ' Red
            .Bold = True
        End With
    End With
End Sub


