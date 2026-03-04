Option Explicit

Private pRows As Collection
Private pNumColumnsMax As Long
Private pNumRowsMax As Long

Public Property Get Rows() As Collection
    Set Rows = pRows
End Property

Public Sub Class_Initialize()
    Set pRows = New Collection
End Sub

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

Private Sub Crop()
    ' Crops each DataRowCls instance s.t. every non-empty cell is only just
    ' included.
    Dim lengths() As Long
    ReDim lengths(1 To pRows.count)
    
    Dim i As Long
    Dim dataRow As DataRowCls
    For i = 1 To pRows.count
        Set dataRow = pRows(i)
        lengths(i) = dataRow.GetIndexOfLastNonEmpty()
    Next i

    Dim maxLength As Long
    maxLength = Application.WorksheetFunction.max(lengths)

    For Each dataRow In pRows
        Call dataRow.CropToLength(maxLength, inplace:=True)
    Next dataRow
End Sub

Private Sub ApplyConditionalFormatting()
    Dim dataRow As DataRowCls
    
    For Each dataRow In pRows
        Call dataRow.ApplyConditionalFormatting
    Next dataRow
End Sub

Private Sub ApplyMacros()
    Dim dataRow As DataRowCls
    
    For Each dataRow In pRows
        Call dataRow.ApplyMacros
    Next dataRow
End Sub

' ----- Public methods -----

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

Public Property Get count() As Long
    count = pRows.count
End Property

Public Property Get Keys() As String()
    Dim result() As String
    Dim i As Long
    ReDim result(1 To pRows.count)

    For i = 1 To pRows.count
        result(i) = pRows(i).key
    Next i
    Keys = result
End Property

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
