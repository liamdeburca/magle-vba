'===============================================================================
' [CLASS] DataRowCls
'===============================================================================
' Description:
'   Represents a single row of measurement data from the Data worksheet. This
'   is the core data structure used throughout the project for storing,
'   manipulating, and displaying batch measurement values. Each row contains
'   metadata (step, name, description, unit) and an array of measurement
'   values that can be filtered, sorted, and validated against min/max bounds.
'
'   DataRowCls instances are created by ParsedDataCls when loading data from
'   the worksheet, and are used by plotting, sorting, and formatting modules.
'
' Properties:
'   step : String
'       Process step identifier in format "xx:xx" (e.g., "01:00")
'   name : String
'       Name of the measurement parameter
'   desc : String
'       Optional description providing additional context
'   unit : String
'       Unit of measurement (e.g., "hh:mm", "txt", or numeric units)
'   target : Variant
'       Target value for the measurement (xlErrNA if not defined)
'   min : Variant
'       Minimum acceptable value (xlErrNA if not defined)
'   max : Variant
'       Maximum acceptable value (xlErrNA if not defined)
'   rowIdx : Long
'       Row index in the Data worksheet
'   RawData : Variant()
'       Original cell values as read from the worksheet
'   DblData : Variant()
'       Values converted to Double where possible
'   txtData : String()
'       Values as text strings
'   Data : Variant()
'       Values converted based on unit type
'   key : String
'       Unique identifier combining step, name, and description
'===============================================================================
Option Explicit

Private pStep As String
Private pName As String
Private pDesc As String
Private pUnit As String
Private pTarget As Variant
Private pMin As Variant
Private pMax As Variant
Private pRowIdx As Long
Private pRawData() As Variant
Private pDblData() As Variant
Private pTxtData() As String
Private pData() As Variant

'===============================================================================
' [PROPERTY] step
'===============================================================================
' Description:
'   Returns the process step identifier for this data row.
'
' Returns:
'   String
'       The step value in format "xx:xx"
'===============================================================================
Public Property Get step() As String
    step = pStep
End Property

'===============================================================================
' [PROPERTY] name
'===============================================================================
' Description:
'   Returns the parameter name for this data row.
'
' Returns:
'   String
'       The measurement parameter name
'===============================================================================
Public Property Get name() As String
    name = pName
End Property

'===============================================================================
' [PROPERTY] desc
'===============================================================================
' Description:
'   Returns the optional description for this data row.
'
' Returns:
'   String
'       Additional context or empty string if not defined
'===============================================================================
Public Property Get desc() As String
    desc = pDesc
End Property

'===============================================================================
' [PROPERTY] unit
'===============================================================================
' Description:
'   Returns the unit of measurement for this data row.
'
' Returns:
'   String
'       Unit string (e.g., "hh:mm", "txt") or empty for numeric values
'===============================================================================
Public Property Get unit() As String
    unit = pUnit
End Property

'===============================================================================
' [PROPERTY] target
'===============================================================================
' Description:
'   Returns the target value for this measurement.
'
' Returns:
'   Variant
'       Target value or xlErrNA if not defined
'===============================================================================
Public Property Get target() As Variant
    target = pTarget
End Property

'===============================================================================
' [PROPERTY] min
'===============================================================================
' Description:
'   Returns the minimum acceptable value for this measurement.
'
' Returns:
'   Variant
'       Minimum value or xlErrNA if not defined
'===============================================================================
Public Property Get min() As Variant
    min = pMin
End Property

'===============================================================================
' [PROPERTY] max
'===============================================================================
' Description:
'   Returns the maximum acceptable value for this measurement.
'
' Returns:
'   Variant
'       Maximum value or xlErrNA if not defined
'===============================================================================
Public Property Get max() As Variant
    max = pMax
End Property

'===============================================================================
' [PROPERTY] rowIdx
'===============================================================================
' Description:
'   Returns the worksheet row index for this data row.
'
' Returns:
'   Long
'       Row number in Data worksheet, or -1 if not loaded from sheet
'===============================================================================
Public Property Get rowIdx() As Long
    rowIdx = pRowIdx
End Property

'===============================================================================
' [PROPERTY] RawData
'===============================================================================
' Description:
'   Returns the original cell values as read from the worksheet.
'
' Returns:
'   Variant()
'       Array of raw cell values
'===============================================================================
Public Property Get RawData() As Variant()
    RawData = pRawData
End Property

'===============================================================================
' [PROPERTY] DblData
'===============================================================================
' Description:
'   Returns values converted to Double where possible.
'
' Returns:
'   Variant()
'       Array of numeric values (xlErrValue for non-numeric cells)
'===============================================================================
Public Property Get DblData() As Variant()
    DblData = pDblData
End Property

'===============================================================================
' [PROPERTY] txtData
'===============================================================================
' Description:
'   Returns values as displayed text strings.
'
' Returns:
'   String()
'       Array of text representations
'===============================================================================
Public Property Get txtData() As String()
    txtData = pTxtData
End Property

'===============================================================================
' [PROPERTY] Data
'===============================================================================
' Description:
'   Returns values converted based on the unit type. For "hh:mm" units,
'   values are converted to time serial numbers. For dates, values are
'   converted to date serial numbers. Otherwise, numeric values are used.
'
' Returns:
'   Variant()
'       Array of converted values appropriate for the unit type
'===============================================================================
Public Property Get Data() As Variant()
    Data = pData
End Property

'===============================================================================
' [PROPERTY] key
'===============================================================================
' Description:
'   Returns a unique identifier string combining step, name, and description.
'   This key is used for looking up rows in ParsedDataCls collections.
'
' Returns:
'   String
'       Identifier in format "[step] name" or "[step] name (desc)"
'===============================================================================
Public Property Get key() As String
    Dim k As String
    k = "[" & pStep & "] " & pName

    If pDesc <> "" Then
        k = k & " (" & pDesc & ")"
    End If

    key = k
End Property

'===============================================================================
' [SUB] Class_Initialize
'===============================================================================
' Description:
'   Initializes a new DataRowCls instance with default values. Called
'   automatically when a new instance is created with "New DataRowCls".
'===============================================================================
Public Sub Class_Initialize()
    pStep = ""
    pName = ""
    pDesc = ""
    pUnit = ""
    pTarget = CVErr(xlErrNA)
    pMin = CVErr(xlErrNA)
    pMax = CVErr(xlErrNA)
    pRowIdx = -1

    ReDim pRawData(0)
    ReDim pDblData(0)
    ReDim pTxtData(0)
    ReDim pData(0)
End Sub

'===============================================================================
' [SUB] LoadFromSheet
'===============================================================================
' Description:
'   Loads data row values from the Data worksheet at the specified row index.
'   Validates that required fields (step, name) are present and properly
'   formatted, then populates all data arrays and writes the generated key
'   back to the worksheet.
'
' Parameters:
'   rowIdx : Long
'       The row number in the Data worksheet to load from
'
' Notes:
'   - Raises error 1001 if step value is missing
'   - Raises error 1002 if step format is invalid (must be xx:xx)
'   - Raises error 1003 if name value is missing
'===============================================================================
Public Sub LoadFromSheet( _
    rowIdx As Long _
)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()
    
    Dim stepVal As Variant
    Dim nameVal As Variant
    stepVal = ws.Range(Specs.StepColumn & rowIdx).value
    nameVal = ws.Range(Specs.NameColumn & rowIdx).value
    
    If IsEmpty(stepVal) Or stepVal = "" Then
        Err.Raise 1001, "DataRow.LoadFromSheet", "Step value is required"
    ElseIf Not IsValidStepFormat(Trim(CStr(stepVal))) Then
        Dim msg As String
        msg = "Step must be in format xx:xx but got " & Trim(CStr(stepVal))
        Err.Raise 1002, "DataRow.LoadFromSheet", msg
    End If
    
    If IsEmpty(nameVal) Or nameVal = "" Then
        Err.Raise 1003, "DataRow.LoadFromSheet", "Name value is required"
    End If
    
    pRowIdx = rowIdx
    pStep = Trim(CStr(stepVal))
    pName = Trim(CStr(nameVal))
    
    Dim descVal As String
    descVal = ws.Range(Specs.DescColumn & rowIdx).value
    If Not IsEmpty(descVal) And descVal <> "" Then
        pDesc = Trim(CStr(descVal))
    End If
    
    Dim unitVal As String
    unitVal = ws.Range(Specs.UnitColumn & rowIdx).value
    If Not IsEmpty(unitVal) And unitVal <> "" Then
        pUnit = Trim(CStr(unitVal))
    End If
    
    Dim targetVal As Variant
    targetVal = ws.Range(Specs.TargetColumn & rowIdx).value
    If Not IsEmpty(targetVal) And targetVal <> "" And Not IsError(targetVal) Then
        pTarget = CDbl(targetVal)
    End If
    
    Dim minVal As Variant
    minVal = ws.Range(Specs.MinColumn & rowIdx).value
    If Not IsEmpty(minVal) And minVal <> "" And Not IsError(minVal) Then
        pMin = CDbl(minVal)
    End If
    
    Dim maxVal As Variant
    maxVal = ws.Range(Specs.MaxColumn & rowIdx).value
    If Not IsEmpty(maxVal) And maxVal <> "" And Not IsError(maxVal) Then
        pMax = CDbl(maxVal)
    End If
    
    ReDim pRawData(1 To Specs.NumColumns)
    ReDim pDblData(1 To Specs.NumColumns)
    ReDim pTxtData(1 To Specs.NumColumns)
    ReDim pData(1 To Specs.NumColumns)
    
    Dim startCol As Long
    startCol = ws.Range(Specs.DataStartColumn & "1").Column
    
    Dim i As Long
    Dim cellValue As Variant
    For i = 1 To Specs.NumColumns
        cellValue = ws.Cells(rowIdx, startCol + i - 1).value
        pRawData(i) = ws.Cells(rowIdx, startCol + i - 1).value
        pTxtData(i) = ws.Cells(rowIdx, startCol + i - 1).Text

        If IsNumeric(pRawData(i)) Then
            pDblData(i) = CDbl(pRawData(i))
        Else
            pDblData(i) = CVErr(xlErrValue)
        End If
    Next i
    
    Call ConvertDataUsingUnit
    Call WriteKey
End Sub

'===============================================================================
' [SUB] ConvertDataUsingUnit
'===============================================================================
' Description:
'   Converts pData array values based on the unit type. Handles time values
'   (hh:mm), date values (åååå-mm-dd, åååå-mm), text values, and numeric
'   values. Also converts target, min, and max bounds as appropriate.
'===============================================================================
Private Sub ConvertDataUsingUnit()
    Dim i As Long
    Dim txtVal As String

    Select Case pUnit
        Case "hh:mm"
            For i = LBound(pData) To UBound(pData)
                txtVal = pTxtData(i)
                If txtVal = "" Then
                    pData(i) = CVErr(xlErrNA)
                Else
                    pData(i) = CDbl(TimeValue(txtVal))
                End If
            Next i
            
        Case "åååå-mm-dd", "åååå-mm"
            For i = LBound(pData) To UBound(pData)
                txtVal = pTxtData(i)
                If txtVal = "" Then
                    pData(i) = CVErr(xlErrNA)
                Else
                    pData(i) = CDbl(DateValue(txtVal))
                End If
            Next i
            If Not IsError(pTarget) Then pTarget = DateValue(CStr(pTarget))
            If Not IsError(pMin) Then pMin = DateValue(CStr(pMin))
            If Not IsError(pMax) Then pMax = DateValue(CStr(pMax))
            
        Case "txt"
            For i = LBound(pData) To UBound(pData)
                pData(i) = pTxtData(i)
            Next i

        Case Else
            For i = LBound(pData) To UBound(pData)
                pData(i) = pDblData(i)
            Next i
            If Not IsError(pTarget) Then pTarget = CDbl(pTarget)
            If Not IsError(pMin) Then pMin = CDbl(pMin)
            If Not IsError(pMax) Then pMax = CDbl(pMax)
    End Select
End Sub

'===============================================================================
' [SUB] WriteKey
'===============================================================================
' Description:
'   Writes the generated key value to the Key column of the Data worksheet.
'   Only writes if rowIdx has been set (greater than 0).
'===============================================================================
Private Sub WriteKey()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()
    
    If pRowIdx > 0 Then
        ws.Range(Specs.KeyColumn & pRowIdx).value = Me.key
    End If
End Sub

'===============================================================================
' [FUNCTION] GetIndexOfLastNonEmpty
'===============================================================================
' Description:
'   Finds the index of the last non-empty value in the data array. Used for
'   cropping data to remove trailing empty cells.
'
' Returns:
'   Long
'       Index of the last non-empty value, or 0 if all values are empty
'===============================================================================
Public Function GetIndexOfLastNonEmpty() As Long
    Dim i As Long
    For i = UBound(pData) To LBound(pData) Step -1
        If Not IsEmpty(pRawData(i)) And pRawData(i) <> "" Then
            GetIndexOfLastNonEmpty = i
            Exit Function
        End If
    Next i
    GetIndexOfLastNonEmpty = 0
End Function

'===============================================================================
' [SUB] MimicDataRow
'===============================================================================
' Description:
'   Copies all properties and data from another DataRowCls instance into
'   this instance. Creates independent copies of all data arrays.
'
' Parameters:
'   other : DataRowCls
'       The source DataRowCls instance to copy from
'===============================================================================
Public Sub MimicDataRow( _
    other As DataRowCls _
)
    pStep = other.step
    pName = other.name
    pDesc = other.desc
    pUnit = other.unit
    pTarget = other.target
    pMin = other.min
    pMax = other.max

    ReDim pRawData(LBound(other.RawData) To UBound(other.RawData))
    ReDim pDblData(LBound(other.DblData) To UBound(other.DblData))
    ReDim pTxtData(LBound(other.txtData) To UBound(other.txtData))
    ReDim pData(LBound(other.Data) To UBound(other.Data))

    Dim oRawData() As Variant
    oRawData = other.RawData
    
    Dim oDblData() As Variant
    oDblData = other.DblData
    
    Dim oTxtData() As String
    oTxtData = other.txtData
    
    Dim oData() As Variant
    oData = other.Data

    Dim i As Long
    For i = LBound(other.RawData) To UBound(other.RawData)
        pRawData(i) = oRawData(i)
        pDblData(i) = oDblData(i)
        pTxtData(i) = oTxtData(i)
        pData(i) = oData(i)
    Next i
End Sub

'===============================================================================
' [FUNCTION] CropToLength
'===============================================================================
' Description:
'   Reduces all data arrays to the specified length, removing trailing values.
'   Used to trim empty cells at the end of data rows.
'
' Parameters:
'   newLength : Long
'       The desired array length
'   inplace : Boolean, Optional
'       If True, modifies this instance. If False, returns a new instance.
'       Default: True
'
' Returns:
'   DataRowCls
'       Reference to this instance (if inplace) or a new cropped instance
'===============================================================================
Public Function CropToLength( _
    newLength As Long, _
    Optional inplace As Boolean = True _
) As DataRowCls
    If Not inplace Then
        Dim newDataRow As New DataRowCls
        Call newDataRow.MimicDataRow(Me)
        Call newDataRow.CropToLength(newLength, True)
        Set CropToLength = newDataRow
        Exit Function
    End If

    If newLength < LBound(pData) Then
        ReDim pRawData(0)
        ReDim pDblData(0)
        ReDim pTxtData(0)
        ReDim pData(0)
    ElseIf newLength < UBound(pData) Then
        ReDim Preserve pRawData(LBound(pRawData) To newLength)
        ReDim Preserve pDblData(LBound(pDblData) To newLength)
        ReDim Preserve pTxtData(LBound(pTxtData) To newLength)
        ReDim Preserve pData(LBound(pData) To newLength)
    End If

    Set CropToLength = Me
End Function

'===============================================================================
' [FUNCTION] ApplyMask
'===============================================================================
' Description:
'   Filters data arrays to keep only values where mask is True. Used for
'   excluding certain data points from analysis or plotting.
'
' Parameters:
'   mask : Boolean()
'       Boolean array with same bounds as pData. True keeps the value.
'   inplace : Boolean, Optional
'       If True, modifies this instance. If False, returns a new instance.
'       Default: True
'
' Returns:
'   DataRowCls
'       Reference to this instance (if inplace) or a new filtered instance
'
' Notes:
'   - Raises error 1006 if mask length doesn't match data length
'===============================================================================
Public Function ApplyMask( _
    mask() As Boolean, _
    Optional inplace As Boolean = True _
) As DataRowCls
    If Not inplace Then
        Dim newDataRow As New DataRowCls
        Call newDataRow.MimicDataRow(Me)
        Call newDataRow.ApplyMask(mask, True)
        Set ApplyMask = newDataRow
        Exit Function
    End If

    If UBound(mask) <> UBound(pData) Then
        Dim msg As String
        msg = "Mask length (" & UBound(mask) & ") does not match data " & _
              "length (" & UBound(pData) & ")"
        Err.Raise 1006, "DataRow.ApplyMask", msg
    End If

    Dim i As Long
    Dim count As Long
    count = 0
    For i = LBound(mask) To UBound(mask)
        If mask(i) Then count = count + 1
    Next i

    If count = 0 Then
        ReDim pRawData(0)
        ReDim pDblData(0)
        ReDim pTxtData(0)
        ReDim pData(0)
        Set ApplyMask = Me
        Exit Function
    End If

    Dim fRawData() As Variant
    Dim fDblData() As Variant
    Dim fTxtData() As String
    Dim fData() As Variant
    ReDim fRawData(1 To count)
    ReDim fDblData(1 To count)
    ReDim fTxtData(1 To count)
    ReDim fData(1 To count)
    
    Dim j As Long
    j = 1
    For i = LBound(mask) To UBound(mask)
        If mask(i) Then
            fRawData(j) = pRawData(i)
            fDblData(j) = pDblData(i)
            fTxtData(j) = pTxtData(i)
            fData(j) = pData(i)
            j = j + 1
        End If
    Next i
    
    pRawData = fRawData
    pDblData = fDblData
    pTxtData = fTxtData
    pData = fData

    Set ApplyMask = Me
End Function

'===============================================================================
' [FUNCTION] GetMask
'===============================================================================
' Description:
'   Creates a boolean mask indicating which data elements are non-empty.
'   Returns True for elements that have values, False for empty cells.
'
' Returns:
'   Boolean()
'       Boolean array with same bounds as pData
'===============================================================================
Public Function GetMask() As Boolean()
    Dim maskCollection As Collection
    Set maskCollection = New Collection

    Dim i As Long
    For i = LBound(pRawData) To UBound(pRawData)
        If Not IsEmpty(pData(i)) And pData(i) <> "" Then
            maskCollection.Add i
        End If
    Next i

    Dim mask() As Boolean
    ReDim mask(LBound(pData) To UBound(pData))

    For i = LBound(pData) To UBound(pData)
        mask(i) = False
    Next i
    
    Dim idx As Variant
    For Each idx In maskCollection
        mask(CLng(idx)) = True
    Next idx
    
    GetMask = mask
End Function

'===============================================================================
' [SUB] UpdateData
'===============================================================================
' Description:
'   Writes pRawData values back to the Data worksheet. Used after sorting
'   or other modifications to persist changes.
'
' Notes:
'   - Raises error 1008 if rowIdx is not set
'===============================================================================
Public Sub UpdateData()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()
    
    If pRowIdx <= 0 Then
        Err.Raise 1008, _
                  "DataRow.UpdateData", _
                  "Row index not set. Cannot update sheet."
    End If
    
    Dim startCol As Long
    startCol = ws.Range(Specs.DataStartColumn & "1").Column
    
    Dim i As Long
    Dim colIdx As Long
    For i = LBound(pData) To UBound(pData)
        colIdx = startCol + (i - LBound(pData))
        ws.Cells(pRowIdx, colIdx).value = pRawData(i)
    Next i
End Sub

'===============================================================================
' [SUB] ApplyConditionalFormatting
'===============================================================================
' Description:
'   Applies conditional formatting rules to this data row's cells in the
'   worksheet. Values below min are highlighted in blue, values above max
'   in red.
'===============================================================================
Public Sub ApplyConditionalFormatting()
    If (Not IsError(pMin) Or Not IsError(pMax)) Then
        Call ApplyConditionalFormattingToDataRow(Me)
    End If
End Sub

'===============================================================================
' [SUB] ApplyMacros
'===============================================================================
' Description:
'   Applies the macro dropdown menu to this data row's macro cell in the
'   worksheet. Enables user interaction for sorting and plotting.
'===============================================================================
Public Sub ApplyMacros()
    Call ApplyMacrosToDataRow(Me)
End Sub

'===============================================================================
' [FUNCTION] GetArgsort
'===============================================================================
' Description:
'   Returns an array of indices that would sort the data. Empty values are
'   always placed last regardless of sort direction.
'
' Parameters:
'   ascending : Boolean, Optional
'       If True, sorts smallest to largest. If False, largest to smallest.
'       Default: True
'
' Returns:
'   Long()
'       Array of indices representing sorted order
'===============================================================================
Public Function GetArgsort( _
    Optional ascending As Boolean = True _
) As Long()
    Dim N As Long
    N = UBound(pRawData) - LBound(pRawData) + 1
    
    Dim indices() As Long
    ReDim indices(LBound(pRawData) To UBound(pRawData))
    
    Dim i As Long
    For i = LBound(pRawData) To UBound(pRawData)
        indices(i) = i
    Next i
    
    Dim j As Long
    Dim temp As Long
    Dim needsSwap As Boolean
    For i = LBound(pRawData) To UBound(pRawData) - 1
        For j = LBound(pRawData) To UBound(pRawData) - 1 - (i - LBound(pRawData))
            needsSwap = False
            
            Dim val1 As Variant
            Dim val2 As Variant
            val1 = pRawData(indices(j))
            val2 = pRawData(indices(j + 1))
            
            Dim isEmpty1 As Boolean
            Dim isEmpty2 As Boolean
            isEmpty1 = IsEmpty(val1) Or val1 = ""
            isEmpty2 = IsEmpty(val2) Or val2 = ""
            
            If isEmpty1 And Not isEmpty2 Then
                needsSwap = True
            ElseIf Not isEmpty1 And isEmpty2 Then
                needsSwap = False
            ElseIf Not isEmpty1 And Not isEmpty2 Then
                Dim cmp As Integer
                cmp = CompareValues(val1, val2)
                
                If ascending Then
                    needsSwap = (cmp > 0)
                Else
                    needsSwap = (cmp < 0)
                End If
            End If
            
            If needsSwap Then
                temp = indices(j)
                indices(j) = indices(j + 1)
                indices(j + 1) = temp
            End If
        Next j
    Next i
    
    GetArgsort = indices
End Function

'===============================================================================
' [FUNCTION] CompareValues
'===============================================================================
' Description:
'   Compares two variant values for sorting purposes.
'
' Parameters:
'   val1 : Variant
'       First value to compare
'   val2 : Variant
'       Second value to compare
'
' Returns:
'   Integer
'       -1 if val1 < val2, 0 if equal, 1 if val1 > val2
'===============================================================================
Private Function CompareValues( _
    val1 As Variant, _
    val2 As Variant _
) As Integer
    On Error GoTo ErrorHandler
    
    If IsNumeric(val1) And IsNumeric(val2) Then
        CompareValues = Sgn(CDbl(val1) - CDbl(val2))
        Exit Function
    End If
    
    Dim str1 As String
    Dim str2 As String
    str1 = CStr(val1)
    str2 = CStr(val2)
    
    If str1 < str2 Then
        CompareValues = -1
    ElseIf str1 > str2 Then
        CompareValues = 1
    Else
        CompareValues = 0
    End If
    
    Exit Function
ErrorHandler:
    CompareValues = 0
End Function

'===============================================================================
' [FUNCTION] ApplyArgsort
'===============================================================================
' Description:
'   Reorders all data arrays based on the provided indices array. Used to
'   apply a sort order computed by GetArgsort.
'
' Parameters:
'   indices : Long()
'       Array of indices representing the desired order
'   inplace : Boolean, Optional
'       If True, modifies this instance. If False, returns a new instance.
'       Default: True
'
' Returns:
'   DataRowCls
'       Reference to this instance (if inplace) or a new reordered instance
'
' Notes:
'   - Raises error 1007 if indices length doesn't match data length
'===============================================================================
Public Function ApplyArgsort( _
    indices() As Long, _
    Optional inplace As Boolean = True _
) As DataRowCls
    If UBound(indices) - LBound(indices) <> UBound(pData) - LBound(pData) Then
        Dim emsg As String
        emsg = "Indices array length (" & _
               (UBound(indices) - LBound(indices) + 1) & _
               ") does not match data length (" & _
               (UBound(pData) - LBound(pData) + 1) & ")"
        Err.Raise 1007, "DataRow.ApplyArgsort", emsg
    End If

    If Not inplace Then
        Dim newDataRow As New DataRowCls
        Call newDataRow.MimicDataRow(Me)
        Call newDataRow.ApplyArgsort(indices, True)
        Set ApplyArgsort = newDataRow
        Exit Function
    End If
    
    Dim sRawData() As Variant
    Dim sDblData() As Variant
    Dim sTxtData() As String
    Dim sData() As Variant
    ReDim sRawData(LBound(pRawData) To UBound(pRawData))
    ReDim sDblData(LBound(pDblData) To UBound(pDblData))
    ReDim sTxtData(LBound(pTxtData) To UBound(pTxtData))
    ReDim sData(LBound(pData) To UBound(pData))
    
    Dim i As Long
    For i = LBound(pData) To UBound(pData)
        sRawData(i) = pRawData(indices(i))
        sDblData(i) = pDblData(indices(i))
        sTxtData(i) = pTxtData(indices(i))
        sData(i) = pData(indices(i))
    Next i
    
    pRawData = sRawData
    pDblData = sDblData
    pTxtData = sTxtData
    pData = sData

    Set ApplyArgsort = Me
End Function

'===============================================================================
' [FUNCTION] IsSorted
'===============================================================================
' Description:
'   Checks if the data is already sorted in the specified order. Empty values
'   are expected to be at the end.
'
' Parameters:
'   ascending : Boolean, Optional
'       If True, checks for ascending order. If False, descending.
'       Default: True
'
' Returns:
'   Boolean
'       True if data is sorted in the specified order, False otherwise
'===============================================================================
Public Function IsSorted( _
    Optional ascending As Boolean = True _
) As Boolean
    Dim i As Long
    Dim lastNonEmptyIdx As Long
    Dim foundNonEmptyYet As Boolean

    foundNonEmptyYet = False
    lastNonEmptyIdx = -1
    
    For i = LBound(pRawData) To UBound(pRawData)
        If Not (IsEmpty(pRawData(i)) Or pRawData(i) = "") Then
            If foundNonEmptyYet And lastNonEmptyIdx >= 0 Then
                Dim nextEmptyIdx As Long
                For nextEmptyIdx = lastNonEmptyIdx + 1 To i - 1
                    If IsEmpty(pRawData(nextEmptyIdx)) _
                        Or (pRawData(nextEmptyIdx) = "") Then
                        IsSorted = False
                        Exit Function
                    End If
                Next nextEmptyIdx
            End If
            
            If foundNonEmptyYet And lastNonEmptyIdx >= 0 Then
                Dim cmp As Integer
                If Not (IsEmpty(pRawData(lastNonEmptyIdx)) _
                    Or pRawData(lastNonEmptyIdx) = "") Then
                    cmp = CompareValues(pRawData(lastNonEmptyIdx), pRawData(i))
                    
                    If ascending Then
                        If cmp > 0 Then
                            IsSorted = False
                            Exit Function
                        End If
                    Else
                        If cmp < 0 Then
                            IsSorted = False
                            Exit Function
                        End If
                    End If
                End If
            End If
            
            foundNonEmptyYet = True
            lastNonEmptyIdx = i
        End If
    Next i
    
    IsSorted = True
End Function

'===============================================================================
' [SUB] Describe
'===============================================================================
' Description:
'   Displays a summary of this DataRowCls instance including key, number of
'   entries, missing values, and bound violations. Outputs to both the
'   Immediate window and a message box.
'===============================================================================
Public Sub Describe()
    Dim msg As String
    msg = Me.key & vbCrLf
    msg = msg & String(Len(Me.key), "-") & vbCrLf
    
    Dim nEntries As Integer
    nEntries = CInt(UBound(Me.Data) - LBound(Me.Data) + 1)
    msg = msg & "Number of entries: " & nEntries & vbCrLf
    
    Dim nMissing As Integer
    nMissing = 0
    
    Dim txtDataArr() As String
    txtDataArr = Me.txtData
    
    Dim i As Long
    For i = LBound(txtDataArr) To UBound(txtDataArr)
        nMissing = nMissing - CInt(txtDataArr(i) = "")
    Next i
    
    msg = msg & "> Missing: " & nMissing & "/" & nEntries & vbCrLf
    msg = msg & "Bounds on values: " & vbCrLf
    
    msg = msg & "> Min: "
    If IsError(Me.min) Then
        msg = msg & "-inf (0)" & vbCrLf
    Else
        msg = msg & Me.min & " (" & _
              StatUtils.CountBelowCritical(Me.DblData, Me.min) & ")" & vbCrLf
    End If
    
    If Not IsError(Me.target) Then
        msg = msg & "> Target: " & Me.target & vbCrLf
    End If
    
    msg = msg & "> Max: "
    If IsError(Me.max) Then
        msg = msg & "inf (0)" & vbCrLf
    Else
        msg = msg & Me.max & " (" & _
              StatUtils.CountAboveCritical(Me.DblData, Me.max) & ")" & vbCrLf
    End If
    
    Debug.Print msg
    MsgBox msg, vbInformation, "(DataRowCls) Describe"
End Sub

'===============================================================================
' [FUNCTION] IsNumeric
'===============================================================================
' Description:
'   Checks if a variant value can be converted to a numeric type.
'
' Parameters:
'   value : Variant
'       The value to check
'
' Returns:
'   Boolean
'       True if value is numeric, False otherwise
'===============================================================================
Private Function IsNumeric( _
    value As Variant _
) As Boolean
    On Error Resume Next
    IsNumeric = Not IsError(CDbl(value))
    On Error GoTo 0
End Function

