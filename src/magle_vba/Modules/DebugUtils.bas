'===============================================================================
' Module: DebugUtils
'===============================================================================
' Description:
'   Debugging utilities for inspecting ParsedDataCls and DataRowCls instances.
'   Provides functions to print summary information and detailed data to the
'   Immediate window and message boxes for troubleshooting.
'===============================================================================

'===============================================================================
' [MACRO] PrintParsedDataInfo
'===============================================================================
' Description:
'   Displays a summary of the current ParsedDataCls instance including the
'   total number of rows loaded and details for the first 10 rows.
'===============================================================================
Public Sub PrintParsedDataInfo()
    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()
    
    Dim output As String
    output = "=== ParsedDataCls INSTANCE INFO ===" & vbCrLf & vbCrLf
    output = output & "Total Rows Loaded: " & ParsedData.count & vbCrLf & vbCrLf
    
    Dim MaxRows As Long
    MaxRows = IIf(ParsedData.count > 10, 10, ParsedData.count)
    
    output = output & "First " & MaxRows & " Rows:" & vbCrLf
    output = output & "---" & vbCrLf
    
    Dim i As Long
    For i = 1 To MaxRows
        Dim dataRow As DataRowCls
        Set dataRow = ParsedData.Rows(i)
        
        Dim dataArray As Variant
        dataArray = dataRow.Data
        
        output = output & "DataRow " & i & ":" & vbCrLf
        output = output & "  Key: " & dataRow.key & vbCrLf
        output = output & "  Step: " & dataRow.step & vbCrLf
        output = output & "  Name: " & dataRow.name & vbCrLf
        output = output & "  Desc: " & dataRow.desc & vbCrLf
        output = output & "  Unit: " & dataRow.unit & vbCrLf
        output = output & "  Data Points: " & UBound(dataArray) & vbCrLf
        output = output & vbCrLf
    Next i
    
    Debug.Print output
    MsgBox output, vbInformation, "Data Instance Debug Info"
End Sub

'===============================================================================
' [SUB] PrintDataRowDetails
'===============================================================================
' Description:
'   Displays detailed information about a specific DataRowCls instance
'   including all metadata, bounds, and data values.
'
' Parameters:
'   rowKey : String
'       The key identifying the data row to inspect
'===============================================================================
Public Sub PrintDataRowDetails( _
    rowKey As String _
)
    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()
    
    Dim Row As DataRowCls
    Set Row = ParsedData.GetRowFromKey(rowKey)
    
    Dim output As String
    output = "=== DataRowCls DETAILS ===" & vbCrLf & vbCrLf
    output = output & "Key: " & Row.key & vbCrLf
    output = output & "Step: " & Row.step & vbCrLf
    output = output & "Name: " & Row.name & vbCrLf
    output = output & "Desc: " & Row.desc & vbCrLf
    output = output & "Unit: " & Row.unit & vbCrLf
    output = output & "Target: " & Row.target & vbCrLf
    output = output & "Min: " & Row.min & vbCrLf
    output = output & "Max: " & Row.max & vbCrLf
    output = output & vbCrLf
    
    Dim dataArray As Variant
    dataArray = Row.Data
    
    output = output & "Data Array (" & UBound(dataArray) & " points):" & vbCrLf
    output = output & "---" & vbCrLf
    
    Dim i As Long
    For i = LBound(dataArray) To UBound(dataArray)
        output = output & "  [" & i & "] = " & dataArray(i) & vbCrLf
    Next i
    
    Debug.Print output
    MsgBox output, vbInformation, "DataRowCls Debug Details"
End Sub