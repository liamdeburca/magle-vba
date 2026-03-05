'===============================================================================
' Module: FormattingUtils
'===============================================================================
' Description:
'   Utilities for conditional formatting of values in the Data sheet.
'   Provides functions to check if values exceed specification bounds and
'   applies visual formatting rules to highlight out-of-spec data points.
'===============================================================================

'===============================================================================
' [FUNCTION] IsBelowMin
'===============================================================================
' Description:
'   Checks whether the value in a cell falls below the minimum specification
'   limit for its corresponding data row.
'
' Parameters:
'   cell : Range
'       The cell to evaluate
'
' Returns:
'   Boolean - True if cell value is below min, False otherwise
'===============================================================================
Public Function IsBelowMin( _
    cell As Range _
) As Boolean
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
'   Checks whether the value in a cell exceeds the maximum specification
'   limit for its corresponding data row.
'
' Parameters:
'   cell : Range
'       The cell to evaluate
'
' Returns:
'   Boolean - True if cell value is above max, False otherwise
'===============================================================================
Public Function IsAboveMax( _
    cell As Range _
) As Boolean
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
'   Applies conditional formatting rules to a data row that highlight values
'   outside specification limits. Values below min are formatted blue bold,
'   values above max are formatted red bold.
'
' Parameters:
'   dataRow : DataRowCls
'       The data row to apply formatting to
'===============================================================================
Public Sub ApplyConditionalFormattingToDataRow( _
    dataRow As DataRowCls _
)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()

    Dim dataStartCol As Long
    dataStartCol = ws.Range(Specs.DataStartColumn & "1").Column

    Dim firstCell As Range
    Set firstCell = ws.Cells(dataRow.rowIdx, dataStartCol)
    
    Dim firstCellRef As String
    firstCellRef = firstCell.address(False, False)

    Dim dataRange As Range
    Set dataRange = ws.Range( _
        ws.Cells(dataRow.rowIdx, dataStartCol), _
        ws.Cells(dataRow.rowIdx, dataStartCol + Specs.NumColumns - 1) _
    )

    dataRange.FormatConditions.Delete

    With dataRange.FormatConditions.Add( _
        xlExpression, _
        , _
        "=IsBelowMin(" & firstCellRef & ")" _
    )
        With .Font
            .Color = RGB(0, 0, 255)
            .Bold = True
        End With
    End With

    With dataRange.FormatConditions.Add( _
        xlExpression, _
        , _
        "=IsAboveMax(" & firstCellRef & ")" _
    )
        With .Font
            .Color = RGB(255, 0, 0)
            .Bold = True
        End With
    End With
End Sub