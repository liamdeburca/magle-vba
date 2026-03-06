Option Explicit

Private pRows As Collection
Private pNumColumnsMax As Long
Private pNumRowsMax As Long

'===============================================================================
' [PROPERTY] Rows
'===============================================================================
' Description:
'   Returns the internal Collection of all loaded DataRowCls instances.
'   Each item in the collection corresponds to one data row from the Data
'   worksheet.
'
' Returns:
'   Collection
'       A Collection of DataRowCls objects, one per loaded row
'===============================================================================
Public Property Get Rows() As Collection
    Set Rows = pRows
End Property

'===============================================================================
' [SUB] Class_Initialize
'===============================================================================
' Description:
'   Initializes a new ParsedDataCls instance. Called automatically when
'   creating a new instance with "New ParsedDataCls". Sets up an empty
'   collection to hold future DataRowCls objects.
'===============================================================================
Public Sub Class_Initialize()
    Set pRows = New Collection
End Sub

'===============================================================================
' [SUB] Init
'===============================================================================
' Description:
'   Loads and processes all data from the Data worksheet. This is the main
'   entry point for populating a ParsedDataCls instance after it is created.
'   Sequentially calls PopulateFromData, Crop, ApplyConditionalFormatting,
'   and ApplyMacros. On error, displays a diagnostic message box.
'
' Notes:
'   - Should be called once after "New ParsedDataCls" via ClsUtils.Start
'   - ClsUtils.GetParsedData handles pausing/resuming screen updates
'===============================================================================
Public Sub Init()
    ' Pause recalculation
    Dim msg As String
    msg = "Initialising ParsedDataCls" & vbCrLf
    msg = msg + String(Len(msg), "-") & vbCrLf
    
    On Error GoTo ErrorHandler
    
    msg = msg & "Calling: PopulateFromData" & vbCrLf
    Call PopulateFromData
    msg = msg & "> Success!" & vbCrLf
    
    msg = msg & "Calling: Crop" & vbCrLf
    Call Crop
    msg = msg & "> Success!" & vbCrLf
    
    msg = msg & "Calling: ApplyConditionalFormatting" & vbCrLf
    Call ApplyConditionalFormatting
    msg = msg & "> Success!" & vbCrLf
    
    msg = msg & "Calling: ApplyMacros" & vbCrLf
    Call ApplyMacros
    msg = msg & "> Success!" & vbCrLf
    
    Exit Sub
ErrorHandler:
    Debug.Print msg
    MsgBox msg, vbInformation, "(ParsedDataCls) Init"
End Sub

' ----- Private methods -----

'===============================================================================
' [SUB] PopulateFromData
'===============================================================================
' Description:
'   Reads the Data worksheet row by row and creates a DataRowCls instance
'   for each valid row (i.e. rows where both step and name are present).
'   Stops reading when an empty row is encountered after at least one row
'   has been successfully loaded. Each DataRowCls instance is added to the
'   internal Rows collection keyed by its unique key string.
'
' Notes:
'   - Reads rows from Specs.DataStartRow to Specs.NumRows
'   - Depends on SpecsCls for column/row configuration
'   - Clears the existing collection before re-populating
'===============================================================================
Private Sub PopulateFromData()
    ' Clear current row collection
    Set pRows = New Collection
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()
    
    Dim dataRow As DataRowCls
    Dim step As Variant, name As Variant, desc As Variant
    
    Dim hasAddedRows As Boolean
    hasAddedRows = False
    
    Dim rowIdx As Long
    For rowIdx = Specs.DataStartRow To Specs.NumRows
        ' Check if first three cells are empty
        step = ws.Range(Specs.StepColumn & rowIdx).value
        name = ws.Range(Specs.NameColumn & rowIdx).value
        desc = ws.Range(Specs.DescColumn & rowIdx).value
        
        If hasAddedRows And _
            (IsEmpty(step) Or step = "") And _
            (IsEmpty(name) Or name = "") And _
            (IsEmpty(desc) Or desc = "") Then
            ' If at least one row has been added and first three cells are empty, stop
            Exit For
            
        ElseIf Not (IsEmpty(step) Or step = "") And Not (IsEmpty(name) Or name = "") Then
            ' Create DataRow
            Set dataRow = New DataRowCls
            Call dataRow.LoadFromSheet(rowIdx)
            Call pRows.Add(dataRow, key:=dataRow.key)
            hasAddedRows = True
        End If
    Next rowIdx
End Sub

'===============================================================================
' [SUB] Crop
'===============================================================================
' Description:
'   Trims the data arrays of all loaded DataRowCls instances to remove
'   trailing empty cells. Finds the maximum occupied column index across
'   all non-formula rows, then crops every row to that length.
'
' Notes:
'   - Rows flagged as formula rows (HasFormula = True) are excluded from
'     the length calculation since they can extend infinitely
'   - Operates in-place on all rows in the collection
'===============================================================================
Private Sub Crop()
    ' Crops each DataRowCls instance s.t. every non-empty cell is only just
    ' included.
    
    Dim i As Long, curLength As Long, maxLength As Long
    Dim dataRow As DataRowCls
    
    maxLength = 0
    For i = 1 To pRows.count
        Set dataRow = pRows(i)
        
        ' DataRows may have infinite values if they're calculated using formulas
        If Not dataRow.HasFormula Then
            curLength = dataRow.GetIndexOfLastNonEmpty
            If curLength > maxLength Then maxLength = curLength
        End If
    Next i
        
    For Each dataRow In pRows
        Call dataRow.CropToLength(maxLength, inplace:=True)
    Next dataRow
End Sub

'===============================================================================
' [SUB] ApplyConditionalFormatting
'===============================================================================
' Description:
'   Applies conditional formatting rules to all loaded DataRowCls instances.
'   Rows with defined min/max bounds get colour-coded highlighting: values
'   below the minimum appear in blue, values above the maximum appear in red.
'
' Notes:
'   - Delegates to DataRowCls.ApplyConditionalFormatting for each row
'   - Formatting is written to the Data worksheet
'===============================================================================
Private Sub ApplyConditionalFormatting()
    Dim dataRow As DataRowCls
    
    For Each dataRow In pRows
        Call dataRow.ApplyConditionalFormatting
    Next dataRow
End Sub

'===============================================================================
' [SUB] ApplyMacros
'===============================================================================
' Description:
'   Adds a dropdown macro selector to the macro column cell of every loaded
'   DataRowCls instance in the Data worksheet. This enables users to trigger
'   actions (Describe, Sort, Scatter) directly from the sheet.
'
' Notes:
'   - Delegates to DataRowCls.ApplyMacros for each row
'   - Dropdown options are defined in MacroUtils.ApplyMacrosToDataRow
'===============================================================================
Private Sub ApplyMacros()
    Dim dataRow As DataRowCls
    
    For Each dataRow In pRows
        Call dataRow.ApplyMacros
    Next dataRow
End Sub

' ----- Public methods -----

'===============================================================================
' [FUNCTION] GetRowFromKey
'===============================================================================
' Description:
'   Retrieves a DataRowCls instance by its unique key string. First attempts
'   an exact match; if that fails, falls back to prefix matching. A prefix
'   match succeeds only when exactly one row starts with the given key.
'
' Parameters:
'   key : String
'       The key to search for, e.g. "[01:00] Temperature" or a prefix
'       thereof
'
' Returns:
'   DataRowCls
'       The matching DataRowCls instance
'
' Notes:
'   - Raises error 1004 if no match is found
'   - Raises error 1005 if multiple prefix matches are found
'===============================================================================
Public Function GetRowFromKey(key As String) As DataRowCls
    ' First pass: check for exact matches
    Dim i As Long
    For i = 1 To pRows.count
        If pRows(i).key = key Then
            Set GetRowFromKey = pRows(i)
            Exit Function
        End If
    Next i
    
    ' Second pass: check for partial matches (prefix matching)
    Dim partialMatches As Collection
    Set partialMatches = New Collection
    
    For i = 1 To pRows.count
        If Left(pRows(i).key, Len(key)) = key Then
            partialMatches.Add pRows(i)
        End If
    Next i
    
    ' Handle results
    Dim emsg As String
    If partialMatches.count = 0 Then
        emsg = "No DataRow found matching key: " & key
        Err.Raise 1004, "ParsedDataCls.GetRowFromKey", emsg
    ElseIf partialMatches.count > 1 Then
        emsg = "Multiple DataRows found matching key: " & key
        emsg = emsg & " (" & partialMatches.count & " matches)"
        Err.Raise 1005, "ParsedDataCls.GetRowFromKey", emsg
    Else
        Set GetRowFromKey = partialMatches(1)
        Exit Function
    End If
End Function

'===============================================================================
' [FUNCTION] GetRowFromIndex
'===============================================================================
' Description:
'   Retrieves a DataRowCls instance by its row index in the Data worksheet.
'   Iterates through all loaded rows and returns the first one whose rowIdx
'   matches the given index.
'
' Parameters:
'   index : Long
'       The worksheet row number to look up (1-based)
'
' Returns:
'   DataRowCls
'       The matching DataRowCls instance
'
' Notes:
'   - Raises error 1004 if no row with the given index exists
'===============================================================================
Public Function GetRowFromIndex(index As Long) As DataRowCls
    Dim dataRow As DataRowCls

    For Each dataRow In pRows
        If dataRow.rowIdx = index Then
            Set GetRowFromIndex = dataRow
            Exit Function
        End If
    Next dataRow

    Dim emsg As String
    emsg = "No DataRow found with row index: " & index
    emsg = emsg & ". Maximum row index is: " & pNumRowsMax
    Err.Raise 1004, "ParsedDataCls.GetRowFromIndex", emsg
End Function

'===============================================================================
' [PROPERTY] count
'===============================================================================
' Description:
'   Returns the number of DataRowCls instances currently loaded.
'
' Returns:
'   Long
'       The number of rows in the collection
'===============================================================================
Public Property Get count() As Long
    count = pRows.count
End Property

'===============================================================================
' [PROPERTY] Keys
'===============================================================================
' Description:
'   Returns an array of all key strings from the loaded DataRowCls instances.
'   Useful for iterating over or searching through available rows.
'
' Returns:
'   String()
'       Array of key strings (1-based), one per loaded row
'===============================================================================
Public Property Get Keys() As String()
    Dim result() As String
    Dim i As Long
    ReDim result(1 To pRows.count)

    For i = 1 To pRows.count
        result(i) = pRows(i).key
    Next i
    Keys = result
End Property

'===============================================================================
' [FUNCTION] SortAlongRow
'===============================================================================
' Description:
'   Sorts all loaded DataRowCls instances by the data values in a specified
'   row. The sort order derived from the reference row is applied uniformly
'   to every other row, keeping column data aligned. Empty values in the
'   reference row are always placed last.
'
' Parameters:
'   key : String
'       The key of the DataRowCls instance whose data determines the sort
'       order
'   ascending : Boolean, Optional
'       If True (default), sorts smallest to largest; if False, largest to
'       smallest
'
' Notes:
'   - Does nothing if the target row is already sorted in the requested
'     direction
'   - All rows are written back to the worksheet after sorting
'===============================================================================
Public Function SortAlongRow(key As String, Optional ascending As Boolean = True)
    ' Sorts all rows in the collection based on the data of the specified row
    ' Input: key - string mapping to a DataRowCls instance
    '        ascending - sort direction (default: True for ascending)
    ' If the row is already sorted, does nothing
    Dim sortRow As DataRowCls
    Set sortRow = Me.GetRowFromKey(key)
    
    ' Check if already sorted
    If sortRow.IsSorted(ascending) Then
        Exit Function
    End If
    
    ' Get the sorted indices
    Dim indices() As Long
    indices = sortRow.GetArgsort(ascending)
    
    ' Apply the sort order to all rows
    Dim i As Long
    Dim dataRow As DataRowCls
    For i = 1 To pRows.count
        Set dataRow = pRows(i)
        Call dataRow.ApplyArgsort(indices)
        Call dataRow.UpdateData
    Next i
End Function
