'===============================================================================
' [CLASS] ParsedDataCls
'===============================================================================
' Description:
'   Container class that manages a collection of DataRowCls instances loaded
'   from the Data worksheet. This is the primary interface for accessing and
'   manipulating batch measurement data. ParsedDataCls handles loading rows
'   from the worksheet, cropping to non-empty columns, applying conditional
'   formatting, and coordinating sorts across all rows.
'
'   Typically accessed via the GetParsedData() function in ClsUtils, which
'   maintains a singleton instance for performance.
'
' Properties:
'   Rows : Collection
'       Collection of DataRowCls instances
'   count : Long
'       Number of data rows loaded
'   Keys : String()
'       Array of all row keys for lookup
'===============================================================================
Option Explicit

Private pRows As Collection
Private pNumColumnsMax As Long
Private pNumRowsMax As Long

'===============================================================================
' [PROPERTY] Rows
'===============================================================================
' Description:
'   Returns the collection of DataRowCls instances.
'
' Returns:
'   Collection
'       Collection containing all loaded DataRowCls objects
'===============================================================================
Public Property Get Rows() As Collection
    Set Rows = pRows
End Property

'===============================================================================
' [SUB] Class_Initialize
'===============================================================================
' Description:
'   Initializes a new ParsedDataCls instance with an empty collection.
'===============================================================================
Public Sub Class_Initialize()
    Set pRows = New Collection
End Sub

'===============================================================================
' [SUB] Init
'===============================================================================
' Description:
'   Loads and processes all data rows from the Data worksheet. Calls
'   PopulateFromData, Crop, ApplyConditionalFormatting, and ApplyMacros
'   in sequence.
'===============================================================================
Public Sub Init()
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

'===============================================================================
' [SUB] PopulateFromData
'===============================================================================
' Description:
'   Reads all valid rows from the Data worksheet and creates DataRowCls
'   instances for each. Stops when encountering empty rows after data.
'===============================================================================
Private Sub PopulateFromData()
    Set pRows = New Collection
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()
    
    Dim hasAddedRows As Boolean
    hasAddedRows = False
    
    Dim rowIdx As Long
    For rowIdx = Specs.DataStartRow To Specs.NumRows
        Dim stepVal As Variant
        Dim nameVal As Variant
        Dim descVal As Variant
        stepVal = ws.Range(Specs.StepColumn & rowIdx).value
        nameVal = ws.Range(Specs.NameColumn & rowIdx).value
        descVal = ws.Range(Specs.DescColumn & rowIdx).value
        
        If hasAddedRows And _
            (IsEmpty(stepVal) Or stepVal = "") And _
            (IsEmpty(nameVal) Or nameVal = "") And _
            (IsEmpty(descVal) Or descVal = "") Then
            Exit For
            
        ElseIf Not (IsEmpty(stepVal) Or stepVal = "") And _
               Not (IsEmpty(nameVal) Or nameVal = "") Then
            Dim dataRow As DataRowCls
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
'   Crops all DataRowCls instances to the length of the longest non-empty
'   data sequence, removing trailing empty cells.
'===============================================================================
Private Sub Crop()
    Dim lengths() As Long
    ReDim lengths(1 To pRows.count)
    
    Dim i As Long
    For i = 1 To pRows.count
        Dim dataRow As DataRowCls
        Set dataRow = pRows(i)
        lengths(i) = dataRow.GetIndexOfLastNonEmpty()
    Next i

    Dim maxLength As Long
    maxLength = Application.WorksheetFunction.max(lengths)

    Dim row As DataRowCls
    For Each row In pRows
        Call row.CropToLength(maxLength, inplace:=True)
    Next row
End Sub

'===============================================================================
' [SUB] ApplyConditionalFormatting
'===============================================================================
' Description:
'   Applies conditional formatting to all data rows that have min/max bounds.
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
'   Applies macro dropdown menus to all data rows.
'===============================================================================
Private Sub ApplyMacros()
    Dim dataRow As DataRowCls
    For Each dataRow In pRows
        Call dataRow.ApplyMacros
    Next dataRow
End Sub

'===============================================================================
' [FUNCTION] GetRowFromKey
'===============================================================================
' Description:
'   Finds and returns a DataRowCls instance by its key. First attempts an
'   exact match, then falls back to prefix matching.
'
' Parameters:
'   key : String
'       The key or key prefix to search for
'
' Returns:
'   DataRowCls
'       The matching DataRowCls instance
'
' Notes:
'   - Raises error 1004 if no matching row is found
'   - Raises error 1005 if multiple rows match a prefix
'===============================================================================
Public Function GetRowFromKey( _
    key As String _
) As DataRowCls
    Dim i As Long
    For i = 1 To pRows.count
        If pRows(i).key = key Then
            Set GetRowFromKey = pRows(i)
            Exit Function
        End If
    Next i
    
    Dim partialMatches As Collection
    Set partialMatches = New Collection
    
    For i = 1 To pRows.count
        If Left(pRows(i).key, Len(key)) = key Then
            partialMatches.Add pRows(i)
        End If
    Next i
    
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
'   Finds and returns a DataRowCls instance by its worksheet row index.
'
' Parameters:
'   index : Long
'       The row number in the Data worksheet
'
' Returns:
'   DataRowCls
'       The matching DataRowCls instance
'
' Notes:
'   - Raises error 1004 if no row exists at the specified index
'===============================================================================
Public Function GetRowFromIndex( _
    index As Long _
) As DataRowCls
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
'   Returns the number of DataRowCls instances in the collection.
'
' Returns:
'   Long
'       Count of loaded data rows
'===============================================================================
Public Property Get count() As Long
    count = pRows.count
End Property

'===============================================================================
' [PROPERTY] Keys
'===============================================================================
' Description:
'   Returns an array of all row keys for enumeration or lookup.
'
' Returns:
'   String()
'       Array of key strings
'===============================================================================
Public Property Get Keys() As String()
    Dim result() As String
    ReDim result(1 To pRows.count)

    Dim i As Long
    For i = 1 To pRows.count
        result(i) = pRows(i).key
    Next i
    Keys = result
End Property

'===============================================================================
' [SUB] SortAlongRow
'===============================================================================
' Description:
'   Sorts all rows in the collection based on the data values of the
'   specified row. This maintains alignment across all measurements when
'   reordering by a particular parameter (e.g., sorting all data by date).
'
' Parameters:
'   key : String
'       Key of the row to use as the sort basis
'   ascending : Boolean, Optional
'       Sort direction. Default: True (ascending)
'===============================================================================
Public Sub SortAlongRow( _
    key As String, _
    Optional ascending As Boolean = True _
)
    Dim sortRow As DataRowCls
    Set sortRow = Me.GetRowFromKey(key)
    
    If sortRow.IsSorted(ascending) Then
        Exit Sub
    End If
    
    Dim indices() As Long
    indices = sortRow.GetArgsort(ascending)
    
    Dim i As Long
    For i = 1 To pRows.count
        Dim dataRow As DataRowCls
        Set dataRow = pRows(i)
        Call dataRow.ApplyArgsort(indices)
        Call dataRow.UpdateData
    Next i
End Sub