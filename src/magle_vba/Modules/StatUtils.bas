Option Explicit

'===============================================================================
' Module: StatUtils
'===============================================================================
' Description:
'   Statistical functions for working with arrays that may contain NA errors.
'   Provides sorting utilities (ArgSort, Sort), descriptive statistics
'   (Mean, Std), quantile/percentile calculations, and counting functions.
'===============================================================================

'===============================================================================
' [FUNCTION] RemoveNA
'===============================================================================
' Description:
'   Filters out NA error values from an array, returning only valid numeric
'   values as a Double array.
'
' Parameters:
'   arr : Variant()
'       Array potentially containing either double or #N/A errors
'
' Returns:
'   Double() - New array with NA values removed
'===============================================================================
Function RemoveNA( _
    arr() As Variant _
) As Double()
    Dim result() As Double
    Dim count As Long
    Dim i As Long
    
    count = 0
    
    For i = LBound(arr) To UBound(arr)
        If Not IsError(arr(i)) Then count = count + 1
    Next i
    
    If count = 0 Then
        ReDim result(1 To 1)
        result(1) = CVErr(xlErrNA)
        RemoveNA = result
        Exit Function
    End If
    
    ReDim result(1 To count)
    count = 1
    
    For i = LBound(arr) To UBound(arr)
        If Not IsError(arr(i)) Then
            result(count) = CDbl(arr(i))
            count = count + 1
        End If
    Next i
    
    RemoveNA = result
End Function

'===============================================================================
' [FUNCTION] ArgSort
'===============================================================================
' Description:
'   Returns an array of indices that would sort the input array.
'   Uses quicksort algorithm internally.
'
' Parameters:
'   arr : Double()
'       Array to compute sorted indices for
'
' Returns:
'   Long() - Array of indices representing sorted order
'===============================================================================
Function ArgSort( _
    arr() As Double _
) As Long()
    Dim indices() As Long
    Dim N As Long
    
    N = UBound(arr) - LBound(arr) + 1
    ReDim indices(1 To N)
    
    Dim i As Long
    For i = 1 To N
        indices(i) = LBound(arr) + i - 1
    Next i
    
    Call QuickSortIndices(arr, indices, 1, N)
    
    ArgSort = indices
End Function

'===============================================================================
' [FUNCTION] Sort
'===============================================================================
' Description:
'   Sorts an array using provided or calculated sort indices.
'
' Parameters:
'   arr : Double()
'       Array to sort
'   sortedIndices : Variant (Optional)
'       Pre-computed indices from ArgSort. If not provided, calculated.
'
' Returns:
'   Double() - Sorted array
'===============================================================================
Function Sort( _
    arr() As Double, _
    Optional sortedIndices As Variant _
) As Double()
    Dim result() As Double
    Dim N As Long
    Dim i As Long
    Dim indices() As Long
    Dim hasIndices As Boolean
    
    N = UBound(arr) - LBound(arr) + 1
    ReDim result(1 To N)
    
    hasIndices = False
    On Error Resume Next
    hasIndices = UBound(sortedIndices) >= LBound(sortedIndices)
    On Error GoTo 0
    
    If Not hasIndices Then
        indices = ArgSort(arr)
    Else
        indices = sortedIndices
    End If
    
    For i = 1 To N
        result(i) = arr(indices(i))
    Next i
    
    Sort = result
End Function

Private Sub QuickSortIndices( _
    arr() As Double, _
    indices() As Long, _
    low As Long, _
    high As Long _
)
    Dim pIndices As Long
    
    If low < high Then
        pIndices = PartitionIndices(arr, indices, low, high)
        Call QuickSortIndices(arr, indices, low, pIndices - 1)
        Call QuickSortIndices(arr, indices, pIndices + 1, high)
    End If
End Sub

Private Function PartitionIndices( _
    arr() As Double, _
    indices() As Long, _
    low As Long, _
    high As Long _
) As Long
    Dim pivot As Double
    Dim i As Long
    Dim j As Long
    Dim temp As Long
    
    pivot = arr(indices(high))
    i = low - 1
    
    For j = low To high - 1
        If arr(indices(j)) < pivot Then
            i = i + 1
            temp = indices(i)
            indices(i) = indices(j)
            indices(j) = temp
        End If
    Next j
    
    temp = indices(i + 1)
    indices(i + 1) = indices(high)
    indices(high) = temp
    
    PartitionIndices = i + 1
End Function

'===============================================================================
' [FUNCTION] Mean
'===============================================================================
' Description:
'   Calculates the arithmetic mean of an array.
'
' Parameters:
'   arr : Double()
'       Array of numeric values
'
' Returns:
'   Variant - Mean value, or #N/A error if array is empty
'===============================================================================
Function Mean( _
    arr() As Double _
) As Variant
    Dim sum As Double
    Dim N As Long
    Dim i As Long
    
    N = UBound(arr) - LBound(arr) + 1
    If N = 0 Then
        Mean = CVErr(xlErrNA)
        Exit Function
    End If
    sum = 0
    
    For i = LBound(arr) To UBound(arr)
        sum = sum + arr(i)
    Next i
    
    Mean = sum / N
End Function

'===============================================================================
' [FUNCTION] Std
'===============================================================================
' Description:
'   Calculates the standard deviation of an array.
'
' Parameters:
'   arr : Double()
'       Array of numeric values
'   ddof : Long (Optional)
'       Delta degrees of freedom. Default 0 for population std dev.
'       Use 1 for sample std dev.
'
' Returns:
'   Variant - Standard deviation, or #N/A error if invalid
'===============================================================================
Function Std( _
    arr() As Double, _
    Optional ddof As Long = 0 _
) As Variant
    Dim M As Double
    Dim sumSquaredDiff As Double
    Dim N As Long
    Dim i As Long
    Dim diff As Double
    
    N = UBound(arr) - LBound(arr) + 1
    If N = 0 Or N = ddof Then
        Std = CVErr(xlErrNA)
        Exit Function
    End If

    M = Mean(arr)
    sumSquaredDiff = 0
    
    For i = LBound(arr) To UBound(arr)
        diff = arr(i) - M
        sumSquaredDiff = sumSquaredDiff + (diff * diff)
    Next i
    
    Std = Sqr(sumSquaredDiff / (N - ddof))
End Function

'===============================================================================
' [FUNCTION] Quantile
'===============================================================================
' Description:
'   Calculates a single quantile of an array using Type 7 linear
'   interpolation (R default).
'
' Parameters:
'   arr : Double()
'       Array of numeric values
'   q : Double
'       Quantile between 0 and 1
'   sortedIndices : Variant (Optional)
'       Pre-computed sort indices to avoid redundant sorting
'
' Returns:
'   Variant - Quantile value, or error if invalid input
'===============================================================================
Function Quantile( _
    arr() As Double, _
    q As Double, _
    Optional sortedIndices As Variant _
) As Variant
    Dim sorted() As Double
    Dim N As Long
    Dim h As Double
    Dim hLower As Long
    Dim hUpper As Long
    Dim result As Double
    
    If q < 0 Or q > 1 Then
        Quantile = CVErr(xlErrNum)
        Exit Function
    End If
    
    N = UBound(arr) - LBound(arr) + 1
    If N = 0 Then
        Quantile = CVErr(xlErrNA)
        Exit Function
    End If

    On Error Resume Next
    If UBound(sortedIndices) >= LBound(sortedIndices) Then
        sorted = Sort(arr, sortedIndices)
    Else
        sorted = Sort(arr)
    End If
    On Error GoTo 0
    
    h = (N - 1) * q + 1
    hLower = Int(h)
    hUpper = hLower + 1
    
    If hLower < LBound(sorted) Then
        result = sorted(LBound(sorted))
    ElseIf hUpper > UBound(sorted) Then
        result = sorted(UBound(sorted))
    Else
        result = sorted(hLower) + (h - hLower) * (sorted(hUpper) - sorted(hLower))
    End If
    
    Quantile = result
End Function

'===============================================================================
' [FUNCTION] Quantiles
'===============================================================================
' Description:
'   Calculates multiple quantiles of an array.
'
' Parameters:
'   arr : Double()
'       Array of numeric values
'   qs : Double()
'       Array of quantiles between 0 and 1
'   sortedIndices : Variant (Optional)
'       Pre-computed sort indices
'
' Returns:
'   Variant() - Array of quantile values
'===============================================================================
Function Quantiles( _
    arr() As Double, _
    qs() As Double, _
    Optional sortedIndices As Variant _
) As Variant()
    Dim result() As Variant
    Dim i As Long
    
    If IsMissing(sortedIndices) Then sortedIndices = ArgSort(arr)
    
    ReDim result(LBound(qs) To UBound(qs))
    
    For i = LBound(qs) To UBound(qs)
        result(i) = Quantile(arr, qs(i), sortedIndices:=sortedIndices)
    Next i
    
    Quantiles = result
End Function

'===============================================================================
' [FUNCTION] Percentile
'===============================================================================
' Description:
'   Calculates a single percentile of an array. Wrapper around Quantile
'   that converts percentage to decimal.
'
' Parameters:
'   arr : Double()
'       Array of numeric values
'   p : Double
'       Percentile between 0 and 100
'   sortedIndices : Variant (Optional)
'       Pre-computed sort indices
'
' Returns:
'   Variant - Percentile value
'===============================================================================
Function Percentile( _
    arr() As Double, _
    p As Double, _
    Optional sortedIndices As Variant _
) As Variant
    If p < 0 Or p > 100 Then
        Percentile = CVErr(xlErrNum)
    Else
        Percentile = Quantile(arr, p / 100, sortedIndices:=sortedIndices)
    End If
End Function

'===============================================================================
' [FUNCTION] Percentiles
'===============================================================================
' Description:
'   Calculates multiple percentiles of an array. Wrapper around Quantiles
'   that converts percentages to decimals.
'
' Parameters:
'   arr : Double()
'       Array of numeric values
'   ps : Double()
'       Array of percentiles between 0 and 100
'   sortedIndices : Variant (Optional)
'       Pre-computed sort indices
'
' Returns:
'   Variant() - Array of percentile values
'===============================================================================
Function Percentiles( _
    arr() As Double, _
    ps() As Double, _
    Optional sortedIndices As Variant _
) As Variant()
    Dim qs() As Double
    Dim i As Long
    
    ReDim qs(LBound(ps) To UBound(ps))
    
    For i = LBound(ps) To UBound(ps)
        If ps(i) < 0 Or ps(i) > 100 Then
            qs(i) = CVErr(xlErrNum)
        Else
            qs(i) = ps(i) / 100
        End If
    Next i
    
    Percentiles = Quantiles(arr, qs, sortedIndices:=sortedIndices)
End Function

'===============================================================================
' [FUNCTION] CountMissing
'===============================================================================
' Description:
'   Counts the number of missing or error elements in an array.
'
' Parameters:
'   arr : Variant()
'       Array potentially containing errors or empty values
'
' Returns:
'   Integer - Count of missing/error elements
'===============================================================================
Function CountMissing( _
    arr() As Variant _
) As Integer
    Dim count As Integer
    count = 0
    
    Dim element As Variant
    For Each element In arr
        count = count - CInt( _
            IsError(element) Or IsMissing(element) Or CStr(element) = "" _
        )
    Next element
    
    CountMissing = count
End Function

'===============================================================================
' [FUNCTION] CountAboveCritical
'===============================================================================
' Description:
'   Counts elements in an array that exceed a critical value.
'
' Parameters:
'   arr : Variant()
'       Array of values to check
'   criticalValue : Variant
'       Threshold value to compare against
'   inclusive : Boolean (Optional)
'       If True, counts values >= criticalValue. Default False (>).
'
' Returns:
'   Integer - Count of elements above critical value
'===============================================================================
Function CountAboveCritical( _
    arr() As Variant, _
    criticalValue As Variant, _
    Optional inclusive As Boolean = False _
) As Integer
    Dim count As Integer
    count = 0
    
    If IsError(criticalValue) Then
        CountAboveCritical = 0
        Exit Function
    End If
    
    Dim element As Variant
    For Each element In arr
        If Not IsError(element) Then
            If inclusive Then
                count = count - CInt(CInt(CDbl(element) >= CDbl(criticalValue)))
            Else
                count = count - CInt(CInt(CDbl(element) > CDbl(criticalValue)))
            End If
        End If
    Next element
    
    CountAboveCritical = count
End Function

'===============================================================================
' [FUNCTION] CountBelowCritical
'===============================================================================
' Description:
'   Counts elements in an array that are below a critical value.
'
' Parameters:
'   arr : Variant()
'       Array of values to check
'   criticalValue : Variant
'       Threshold value to compare against
'   inclusive : Boolean (Optional)
'       If True, counts values <= criticalValue. Default False (<).
'
' Returns:
'   Integer - Count of elements below critical value
'===============================================================================
Function CountBelowCritical( _
    arr() As Variant, _
    criticalValue As Variant, _
    Optional inclusive As Boolean = False _
) As Integer
    Dim count As Integer
    count = 0
    
    If IsError(criticalValue) Then
        CountBelowCritical = 0
        Exit Function
    End If
    
    Dim element As Variant
    For Each element In arr
        If Not IsError(element) Then
            If inclusive Then
                count = count - CInt(CInt(CDbl(element) <= CDbl(criticalValue)))
            Else
                count = count - CInt(CInt(CDbl(element) < CDbl(criticalValue)))
            End If
        End If
    Next element
    
    CountBelowCritical = count
End Function
