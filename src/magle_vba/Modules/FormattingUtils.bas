' Utilities For formatting values in the Data sheet

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
    firstCellRef = firstCell.address(False, False)  ' Relative reference like A5

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


