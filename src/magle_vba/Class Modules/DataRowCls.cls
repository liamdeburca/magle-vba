' DataRow Class - Represents a data row with step information and measurements
Option Explicit

Private pStep As String
Private pName As String
Private pDesc As String
Private pUnit As String

Private pTarget As Variant
Private pMin  As Variant
Private pMax  As Variant

Private pRowIdx As Long

Private pRawData() As Variant
Private pDblData() As Variant
Private pTxtData() As String
Private pData() As Variant

Public Property Get step() As String
    step = pStep
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Get desc() As String
    desc = pDesc
End Property

Public Property Get unit() As String
    unit = pUnit
End Property

Public Property Get target() As Variant
    target = pTarget
End Property

Public Property Get min() As Variant
    min = pMin
End Property

Public Property Get max() As Variant
    max = pMax
End Property

Public Property Get rowIdx() As Long
    rowIdx = pRowIdx
End Property

Public Property Get RawData() As Variant()
    RawData = pRawData
End Property

Public Property Get DblData() As Variant()
    DblData = pDblData
End Property

Public Property Get txtData() As String()
    txtData = pTxtData
End Property

Public Property Get Data() As Variant()
    Data = pData
End Property

Public Property Get key() As String
    Dim k As String
    k = "[" & pStep & "] " & pName

    If pDesc <> "" Then
        k = k & " (" & pDesc & ")"
    End If

    key = k
End Property

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

'' ===== LOAD METHOD =====
Public Sub LoadFromSheet(rowIdx As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()
    
    Dim step As Variant
    Dim name As Variant

    step = ws.Range(Specs.StepColumn & rowIdx).value
    name = ws.Range(Specs.NameColumn & rowIdx).value
    
    ' Validate that step and name are defined
    If IsEmpty(step) Or step = "" Then
        Err.Raise 1001, "DataRow.LoadFromSheet", "Step value is required"
    ElseIf Not IsValidStepFormat(Trim(CStr(step))) Then
        Dim msg As String
        msg = "Step must be in format xx:xx but got " & Trim(CStr(step))
        Err.Raise 1002, "DataRow.LoadFromSheet", msg
    End If
    
    If IsEmpty(name) Or name = "" Then
        Err.Raise 1003, "DataRow.LoadFromSheet", "Name value is required"
    End If
    
    ' All validations passed, now load the row
    pRowIdx = rowIdx
    pStep = Trim(CStr(step))
    pName = Trim(CStr(name))
    
    ' Load optional fields only if not empty
    Dim desc As String, unit As String
    Dim target As Variant, min As Variant, max As Variant
    
    desc = ws.Range(Specs.DescColumn & rowIdx).value
    If Not IsEmpty(desc) And desc <> "" Then
        pDesc = Trim(CStr(desc))
    End If
    
    unit = ws.Range(Specs.UnitColumn & rowIdx).value
    If Not IsEmpty(unit) And unit <> "" Then
        pUnit = Trim(CStr(unit))
    End If
    
    target = ws.Range(Specs.TargetColumn & rowIdx).value
    If Not IsEmpty(target) And target <> "" And Not IsError(target) Then
        pTarget = CDbl(target)
    End If
    
    min = ws.Range(Specs.MinColumn & rowIdx).value
    If Not IsEmpty(min) And min <> "" And Not IsError(min) Then
        pMin = CDbl(min)
    End If
    
    max = ws.Range(Specs.MaxColumn & rowIdx).value
    If Not IsEmpty(max) And max <> "" And Not IsError(max) Then
        pMax = CDbl(max)
    End If
    
    Dim i As Long
    ReDim pRawData(1 To Specs.NumColumns)
    ReDim pDblData(1 To Specs.NumColumns)
    ReDim pTxtData(1 To Specs.NumColumns)
    ReDim pData(1 To Specs.NumColumns)
    
    ' Convert column letter to number
    Dim startCol As Long
    startCol = ws.Range(Specs.DataStartColumn & "1").Column
    
    ' Collect data from DataStartColumn for NumColumns columns
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

' ----- METHODS TO PROCESS DATA AFTER LOADING ------

Private Sub ConvertDataUsingUnit()
    Dim i As Long
    Dim txtVal As String

    ' Determine data type based on unit
    Select Case pUnit
        Case "hh:mm"
            ' Convert to time (stored as double with time serial number)
            For i = LBound(pData) To UBound(pData)
                txtVal = pTxtData(i)
                If txtVal = "" Then
                    pData(i) = CVErr(xlErrNA)
                Else
                    pData(i) = CDbl(TimeValue(txtVal))
                End If
            Next i
            ' Convert target, min, max
            '' Need to implement target/min/max conversions when needed
            
        Case "åååå-mm-dd", "åååå-mm"
            ' Convert to date
            For i = LBound(pData) To UBound(pData)
                txtVal = pTxtData(i)
                If txtVal = "" Then
                    pData(i) = CVErr(xlErrNA)
                Else
                    pData(i) = CDbl(DateValue(txtVal))
                End If
            Next i
            ' Convert target, min, max
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
            ' Convert target, min, max
            If Not IsError(pTarget) Then pTarget = CDbl(pTarget)
            If Not IsError(pMin) Then pMin = CDbl(pMin)
            If Not IsError(pMax) Then pMax = CDbl(pMax)
    End Select
End Sub

Private Sub WriteKey()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()
    
    If pRowIdx > 0 Then
        ws.Range(Specs.KeyColumn & pRowIdx).value = Me.key
    End If
End Sub

Public Function GetIndexOfLastNonEmpty() As Long
    Dim i As Long
    For i = UBound(pData) To LBound(pData) Step -1
        If Not IsEmpty(pRawData(i)) And pRawData(i) <> "" Then
            GetIndexOfLastNonEmpty = i
            Exit Function
        End If
    Next i
    GetIndexOfLastNonEmpty = 0 ' All values are empty
End Function

Public Sub MimicDataRow(other As DataRowCls)
    pStep = other.step
    pName = other.name
    pDesc = other.desc
    pUnit = other.unit
    pTarget = other.target
    pMin = other.min
    pMax = other.max

    ' Create independent copies of data arrays
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

Public Function CropToLength( _
    newLength As Long, _
    Optional inplace As Boolean = True _
) As DataRowCls
    ' Crops the pData array to the specified length

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

Public Function ApplyMask( _
    mask() As Boolean, _
    Optional inplace As Boolean = True _
) As DataRowCls
    ' Validate mask and data have compatible dimensions
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

    ' Apply mask - keep only values where mask is True
    Dim fRawData() As Variant
    Dim fDblData() As Variant
    Dim fTxtData() As String
    Dim fData() As Variant

    Dim i As Long, count As Long
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
    
    ' Update original arrays with filtered data
    pRawData = fRawData
    pDblData = fDblData
    pTxtData = fTxtData
    pData = fData

    Set ApplyMask = Me
End Function

Public Function GetMask() As Boolean()
    Dim maskTemp As Collection

    Dim maskCollection As Collection
    Set maskCollection = New Collection

    Dim i As Long
    For i = LBound(pRawData) To UBound(pRawData)
        If Not IsEmpty(pData(i)) And pData(i) <> "" Then maskCollection.Add i
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

Public Sub UpdateData()
    ' Writes pRawData values back to the sheet
    ' Uses SpecsCls to determine which columns and rows to write to
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()
    
    ' Validate that row index is set
    If pRowIdx <= 0 Then
        Err.Raise 1008, _
                  "DataRow.UpdateData", _
                  "Row index not set. Cannot update sheet."
    End If
    
    ' Convert DataStartColumn letter to column number
    Dim startCol As Long
    startCol = ws.Range(Specs.DataStartColumn & "1").Column
    
    ' Write each data value to the corresponding sheet cell
    Dim i As Long, colIdx As Long
    Dim cell As Range
    For i = LBound(pData) To UBound(pData)
        colIdx = startCol + (i - LBound(pData))
        With ws.Cells(pRowIdx, colIdx)
            .value = pRawData(i)
        End With
    Next i
End Sub

'' ====== Formatting and Sheet I/O methods ======

Public Sub ApplyConditionalFormatting()
    If (Not IsError(pMin) Or Not IsError(pMax)) Then
        Call ApplyConditionalFormattingToDataRow(Me)
    End If
End Sub

Public Sub ApplyMacros()
    Call ApplyMacrosToDataRow(Me)
End Sub

'' ===== SORTING METHODS =====

Public Function GetArgsort(Optional ascending As Boolean = True) As Long()
    ' Returns indices that would sort the pRawData array
    ' Empty values are always placed last
    ' ascending: True for ascending sort, False for descending
    
    Dim N As Long
    Dim indices() As Long
    Dim i As Long, j As Long
    Dim temp As Long
    Dim minIdx As Long, maxIdx As Long
    Dim needsSwap As Boolean
    
    N = UBound(pRawData) - LBound(pRawData) + 1
    
    ' Initialize indices array (1-based, matching pRawData)
    ReDim indices(LBound(pRawData) To UBound(pRawData))
    For i = LBound(pRawData) To UBound(pRawData)
        indices(i) = i
    Next i
    
    ' Bubble sort with custom comparator
    ' Separates empty values to the end
    For i = LBound(pRawData) To UBound(pRawData) - 1
        For j = LBound(pRawData) To UBound(pRawData) - 1 - (i - LBound(pRawData))
            needsSwap = False
            
            ' Get values at current indices
            Dim val1 As Variant, val2 As Variant
            val1 = pRawData(indices(j))
            val2 = pRawData(indices(j + 1))
            
            Dim isEmpty1 As Boolean, isEmpty2 As Boolean
            isEmpty1 = IsEmpty(val1) Or val1 = ""
            isEmpty2 = IsEmpty(val2) Or val2 = ""
            
            ' Rule: Empty values always go to the end
            If isEmpty1 And Not isEmpty2 Then
                needsSwap = True
            ElseIf Not isEmpty1 And isEmpty2 Then
                needsSwap = False
            ElseIf Not isEmpty1 And Not isEmpty2 Then
                ' Both non-empty: compare based on sort direction
                Dim cmp As Integer
                cmp = CompareValues(val1, val2)
                
                If ascending Then
                    needsSwap = (cmp > 0)
                Else
                    needsSwap = (cmp < 0)
                End If
            End If
            
            ' Swap if needed
            If needsSwap Then
                temp = indices(j)
                indices(j) = indices(j + 1)
                indices(j + 1) = temp
            End If
        Next j
    Next i
    
    GetArgsort = indices
End Function

Private Function CompareValues(val1 As Variant, val2 As Variant) As Integer
    ' Compares two variant values
    ' Returns: -1 if val1 < val2, 0 if equal, 1 if val1 > val2
    
    On Error GoTo ErrorHandler
    
    ' Try numeric comparison first
    If IsNumeric(val1) And IsNumeric(val2) Then
        CompareValues = Sgn(CDbl(val1) - CDbl(val2))
        Exit Function
        End If
    
    ' String comparison
    Dim str1 As String, str2 As String
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

Public Function ApplyArgsort( _
    indices() As Long, _
    Optional inplace As Boolean = True _
) As DataRowCls
    ' Reorders pData based on the provided indices array
    ' indices should be same length and base as pData
    
    ' Validate indices array matches pData dimensions
    If UBound(indices) - LBound(indices) <> UBound(pData) - LBound(pData) Then
        Dim emsg As String
        emsg = "Indices array length (" & (UBound(indices) - LBound(indices) + 1) & _
               ") does not match data length (" & (UBound(pData) - LBound(pData) + 1) & ")"
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
    
    ' Apply reordering
    Dim i As Long
    For i = LBound(pData) To UBound(pData)
        sRawData(i) = pRawData(indices(i))
        sDblData(i) = pDblData(indices(i))
        sTxtData(i) = pTxtData(indices(i))
        sData(i) = pData(indices(i))
    Next i
    
    ' Replace pData with reordered data
    pRawData = sRawData
    pDblData = sDblData
    pTxtData = sTxtData
    pData = sData

    Set ApplyArgsort = Me
End Function

Public Function IsSorted(Optional ascending As Boolean = True) As Boolean
    ' Checks if pData is already sorted in the specified order
    ' Returns True if sorted, False otherwise
    ' Empty values are expected to be at the end
    
    Dim i As Long
    Dim lastNonEmptyIdx As Long
    Dim lastNonEmptyVal As Variant
    Dim foundNonEmptyYet As Boolean
    
    foundNonEmptyYet = False
    lastNonEmptyIdx = -1
    
    ' Iterate through all data values
    For i = LBound(pRawData) To UBound(pRawData)
        If Not (IsEmpty(pRawData(i)) Or pRawData(i) = "") Then
            ' Check if we've already found empty values before this non-empty value
            ' That would mean data is not properly sorted (empty should be at end)
            If foundNonEmptyYet And lastNonEmptyIdx >= 0 Then
                Dim nextEmptyIdx As Long
                
                For nextEmptyIdx = lastNonEmptyIdx + 1 To i - 1
                    If IsEmpty(pRawData(nextEmptyIdx)) _
                        Or (pRawData(nextEmptyIdx) = "") Then
                        ' Found empty value before current non-empty value
                        ' This violates sort order
                        IsSorted = False
                        Exit Function
                    End If
                Next nextEmptyIdx
            End If
            
            ' If this is not the first non-empty value, check sort order
            If foundNonEmptyYet And lastNonEmptyIdx >= 0 Then
                Dim cmp As Integer
                If Not (IsEmpty(pRawData(lastNonEmptyIdx)) _
                    Or pRawData(lastNonEmptyIdx) = "") Then
                    cmp = CompareValues(pRawData(lastNonEmptyIdx), pRawData(i))
                    
                    If ascending Then
                        ' For ascending: previous value should be <= current value
                        If cmp > 0 Then
                            IsSorted = False
                            Exit Function
                        End If
                    Else
                        ' For descending: previous value should be >= current value
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

Public Sub Describe()
    ' Describes this DataRowCls instance
    Dim msg As String
    Dim i As Long
    
    msg = Me.key & vbCrLf
    msg = msg & String(Len(Me.key), "-") & vbCrLf
    
    ' Number of entries
    Dim nEntries As Integer
    nEntries = CInt(UBound(Me.Data) - LBound(Me.Data) + 1)
    msg = msg & "Number of entries: " & nEntries & vbCrLf
    
    Dim nMissing As Integer
    Dim txtData() As String
    
    nMissing = 0
    txtData = Me.txtData
    For i = LBound(txtData) To UBound(txtData)
        nMissing = nMissing - CInt(txtData(i) = "")
    Next i
    
    msg = msg & "> Missing: " & nMissing & "/" & nEntries & vbCrLf
    
    ' Bounds on values
    msg = msg & "Bounds on values: " & vbCrLf
    
    msg = msg & "> Min: "
    If IsError(Me.min) Then
        msg = msg & "-inf (0)" & vbCrLf
    Else
        msg = msg & Me.min & " (" & StatUtils.CountBelowCritical(Me.DblData, Me.min) & ")" & vbCrLf
    End If
    
    If Not IsError(Me.target) Then
        msg = msg & "> Target: " & Me.target & vbCrLf
    End If
    
    msg = msg & "> Max: "
    If IsError(Me.max) Then
        msg = msg & "inf (0)" & vbCrLf
    Else
        msg = msg & Me.max & " (" & StatUtils.CountAboveCritical(Me.DblData, Me.max) & ")" & vbCrLf
    End If
    
    Debug.Print msg
    MsgBox msg, vbInformation, "(DataRowCls) Describe"
End Sub

' --------------------------------------------------------------------------
' Utilities
' --------------------------------------------------------------------------

Private Function IsNumeric(value As Variant) As Boolean
    ' Check if value is numeric
    On Error Resume Next
    IsNumeric = Not IsError(CDbl(value))
    On Error GoTo 0
End Function

