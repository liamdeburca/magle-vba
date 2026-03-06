''
'' Debug Module - Convenience functions for debugging ParsedDataCls and DataRowCls instances
''

'===============================================================================
' [MACRO] PrintParsedDataInfo
'===============================================================================
' Description:
'   Prints a summary of the currently loaded ParsedDataCls instance to the
'   Immediate Window and displays it in a message box. Shows the total
'   number of loaded rows and detailed information for up to the first ten
'   rows (key, step, name, description, unit, number of data points).
'
' Notes:
'   - Useful for verifying that data loaded correctly after calling Start
'   - Output is truncated to the first 10 rows to keep the message readable
'===============================================================================
Public Sub PrintParsedDataInfo()
    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()
    
    Dim output As String
    output = "=== ParsedDataCls INSTANCE INFO ===" & vbCrLf & vbCrLf
    
    output = output & "Total Rows Loaded: " & ParsedData.count & vbCrLf & vbCrLf
    
    ' Print first 10 rows
    Dim MaxRows As Long
    MaxRows = IIf(ParsedData.count > 10, 10, ParsedData.count)
    
    output = output & "First " & MaxRows & " Rows:" & vbCrLf
    output = output & "---" & vbCrLf
    
    Dim i As Long
    Dim dataRow As DataRowCls
    Dim dataArray As Variant
    
    For i = 1 To MaxRows
        Set dataRow = ParsedData.Rows(i)
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
    
    ' Print to Immediate Window
    Debug.Print output
    
    ' Also show in message box for visibility
    MsgBox output, vbInformation, "Data Instance Debug Info"
End Sub

'===============================================================================
' [SUB] PrintDataRowDetails
'===============================================================================
' Description:
'   Prints full details for a single DataRowCls instance identified by its
'   key string. Outputs key, step, name, description, unit, goal, min, max,
'   and the full data array to the Immediate Window and a message box.
'
' Parameters:
'   rowKey : String
'       The key of the DataRowCls to inspect, e.g. "[01:00] Temperature"
'
' Notes:
'   - Raises an error if the key is not found in the current ParsedDataCls
'===============================================================================
Public Sub PrintDataRowDetails(rowKey As String)
    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()
    
    Dim dataRow As DataRowCls
    Set Row = ParsedData.GetRowFromKey(rowKey)
    
    Dim output As String
    output = "=== DataRowCls DETAILS ===" & vbCrLf & vbCrLf
    output = output & "Key: " & Row.key & vbCrLf
    output = output & "Step: " & Row.step & vbCrLf
    output = output & "Name: " & Row.name & vbCrLf
    output = output & "Desc: " & Row.desc & vbCrLf
    output = output & "Unit: " & Row.unit & vbCrLf
    output = output & "Goal: " & Row.goal & vbCrLf
    output = output & "Min: " & Row.min & vbCrLf
    output = output & "Max: " & Row.max & vbCrLf
    output = output & vbCrLf
    
    ' Print data array
    Dim dataArray As Variant
    dataArray = dataRow.Data
    
    output = output & "Data Array (" & UBound(dataArray) & " points):" & vbCrLf
    output = output & "---" & vbCrLf
    
    Dim i As Long
    For i = LBound(dataArray) To UBound(dataArray)
        output = output & "  [" & i & "] = " & dataArray(i) & vbCrLf
    Next i
    
    Debug.Print output
    MsgBox output, vbInformation, "DataRowCls Debug Details"
End Sub

'===============================================================================
' [MACRO] Test
'===============================================================================
' Description:
'   Simple smoke-test macro that instantiates a SpecsCls object and displays
'   its NumRows value in a message box. Used to verify that basic VBA class
'   initialisation is working correctly during development.
'===============================================================================
Sub Test()
    Dim Specs As SpecsCls
    Set Specs = ClsUtils.GetSpecs()
    MsgBox Specs.NumRows
End Sub
