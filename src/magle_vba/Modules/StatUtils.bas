Option Explicit

' Statistical functions for working with double arrays containing potential NA errors

' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================

'===============================================================================
' [FUNCTION] RemoveNA
'===============================================================================
' Description:
'   Filters a Variant array by removing elements that are Excel error values
'   (e.g. #N/A, #VALUE!). Returns a clean Double array containing only the
'   non-error elements. Used before calling statistical functions such as
'   Mean or Std to avoid runtime errors.
'
' Parameters:
'   arr : Variant()
'       Array potentially containing error values mixed with numeric values
'
' Returns:
'   Double()
'       New 1-based array containing only the non-error elements as Doubles.
'       If all elements are errors, returns a single-element array containing
'       CVErr(xlErrNA).
'===============================================================================
Function RemoveNA(arr() As Variant) As Double()
    Dim result() As Double
    Dim count As Long
    Dim i As Long
    Dim filtered As Boolean
    
    count = 0
    
    ' Count non-NA values
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
    
    ' Copy non-NA values
    For i = LBound(arr) To UBound(arr)
        If Not IsError(arr(i)) Then
            result(count) = CDbl(arr(i))
            count = count + 1
        End If
    Next i
    
    RemoveNA = result
End Function

' ============================================================================
' SORTING FUNCTIONS
' ============================================================================

'===============================================================================
' [FUNCTION] ArgSort
'===============================================================================
' Description:
'   Returns an array of indices that would sort the input array in ascending
'   order using the QuickSort algorithm. The indices are 1-based and
'   reference positions within the input array.
'
' Parameters:
'   arr : Double()
'       The array of values to sort
'
' Returns:
'   Long()
'       1-based array of indices representing the ascending sorted order
'
' Example:
'   Dim idx() As Long
'   idx = ArgSort(myArray)
'===============================================================================
Function ArgSort(arr() As Double) As Long()
    Dim indices() As Long
    Dim N As Long
    
    N = UBound(arr) - LBound(arr) + 1
    ReDim indices(1 To N)
    
    ' Initialize indices
    Dim i As Long
    For i = 1 To N
        indices(i) = LBound(arr) + i - 1
    Next i
    
    ' QuickSort on indices
    Call QuickSortIndices(arr, indices, 1, N)
    
    ArgSort = indices
End Function

'===============================================================================
' [FUNCTION] Sort
'===============================================================================
' Description:
'   Returns a sorted copy of the input Double array in ascending order.
'   If pre-computed sort indices are supplied (e.g. from ArgSort), they are
'   used directly to avoid recomputing the sort order.
'
' Parameters:
'   arr : Double()
'       The array to sort
'   sortedIndices : Variant, Optional
'       Pre-computed index array from ArgSort. If omitted, ArgSort is
'       called internally.
'
' Returns:
'   Double()
'       New 1-based array containing the sorted values
'
' Example:
'   Dim sorted() As Double
'   sorted = Sort(myArray)
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
    
    ' Check if sortedIndices was provided
    hasIndices = False
    On Error Resume Next
    hasIndices = UBound(sortedIndices) >= LBound(sortedIndices)
    On Error GoTo 0
    
    If Not hasIndices Then
        indices = ArgSort(arr)
    Else
        indices = sortedIndices
    End If
    
    ' Build sorted result
    For i = 1 To N
        result(i) = arr(indices(i))
    Next i
    
    Sort = result
End Function

' Internal helper: QuickSort for indices
'===============================================================================
' [SUB] QuickSortIndices
'===============================================================================
' Description:
'   Recursive QuickSort implementation that sorts an index array in-place
'   by comparing values in the corresponding data array. Used internally
'   by ArgSort.
'
' Parameters:
'   arr : Double()
'       The data array whose values determine sort order
'   indices : Long()
'       The index array to sort in-place
'   low : Long
'       Lower bound of the current partition (1-based)
'   high : Long
'       Upper bound of the current partition (1-based)
'===============================================================================
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

' Internal helper: Partition for QuickSort
'===============================================================================
' [FUNCTION] PartitionIndices
'===============================================================================
' Description:
'   Partitions a sub-range of the index array around a pivot value for
'   use in QuickSortIndices. Elements whose corresponding data values are
'   less than the pivot are moved before it; others after.
'
' Parameters:
'   arr : Double()
'       The data array whose values determine ordering
'   indices : Long()
'       The index array being partitioned in-place
'   low : Long
'       Start of the sub-range to partition
'   high : Long
'       End of the sub-range; the pivot element is at indices(high)
'
' Returns:
'   Long
'       The final position of the pivot element after partitioning
'===============================================================================
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

' ============================================================================
' DESCRIPTIVE STATISTICS
' ============================================================================

'===============================================================================
' [FUNCTION] Mean
'===============================================================================
' Description:
'   Calculates the arithmetic mean (average) of a Double array. Returns
'   xlErrNA if the array is empty.
'
' Parameters:
'   arr : Double()
'       Array of numeric values to average
'
' Returns:
'   Variant
'       The mean as a Double, or CVErr(xlErrNA) if the array is empty
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
'   Calculates the standard deviation of a Double array. Supports both
'   population (ddof=0) and sample (ddof=1) standard deviation via the
'   degrees-of-freedom correction parameter.
'
' Parameters:
'   arr : Double()
'       Array of numeric values
'   ddof : Long, Optional
'       Degrees-of-freedom correction. 0 = population std dev (default),
'       1 = sample std dev. Default: 0
'
' Returns:
'   Variant
'       Standard deviation as a Double, or CVErr(xlErrNA) if the array
'       is empty or has fewer elements than ddof
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

' ============================================================================
' QUANTILE FUNCTIONS
' ============================================================================

'===============================================================================
' [FUNCTION] Quantile
'===============================================================================
' Description:
'   Calculates a single quantile of a Double array using linear
'   interpolation (Type 7, the default method used by R and NumPy). The
'   array is sorted before computation unless pre-sorted indices are
'   provided.
'
' Parameters:
'   arr : Double()
'       Array of numeric values
'   q : Double
'       Quantile to compute, between 0 and 1 (e.g. 0.5 for the median)
'   sortedIndices : Variant, Optional
'       Pre-computed sort indices from ArgSort to avoid resorting
'
' Returns:
'   Variant
'       The quantile value as a Double, or CVErr(xlErrNA) if the array is
'       empty, or CVErr(xlErrNum) if q is outside [0, 1]
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

    ' Check if sortedIndices was provided
    On Error Resume Next
    If UBound(sortedIndices) >= LBound(sortedIndices) Then
        sorted = Sort(arr, sortedIndices)
    Else
        sorted = Sort(arr)
    End If
    On Error GoTo 0
    
    ' Type 7 interpolation: h = (n-1)*q + 1
    h = (N - 1) * q + 1
    hLower = Int(h)
    hUpper = hLower + 1
    
    If hLower < LBound(sorted) Then
        result = sorted(LBound(sorted))
    ElseIf hUpper > UBound(sorted) Then
        result = sorted(UBound(sorted))
    Else
        ' Linear interpolation
        result = sorted(hLower) + (h - hLower) * (sorted(hUpper) - sorted(hLower))
    End If
    
    Quantile = result
End Function

'===============================================================================
' [FUNCTION] Quantiles
'===============================================================================
' Description:
'   Calculates multiple quantiles of a Double array in a single call.
'   Pre-computes sort indices once and reuses them for each quantile,
'   making this more efficient than calling Quantile repeatedly.
'
' Parameters:
'   arr : Double()
'       Array of numeric values
'   qs : Double()
'       Array of quantile levels, each between 0 and 1
'   sortedIndices : Variant, Optional
'       Pre-computed sort indices from ArgSort
'
' Returns:
'   Variant()
'       Array of quantile values with the same bounds as qs
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

' ============================================================================
' PERCENTILE FUNCTIONS (Wrappers for Quantile)
' ============================================================================

'===============================================================================
' [FUNCTION] Percentile
'===============================================================================
' Description:
'   Calculates a single percentile of a Double array. This is a convenience
'   wrapper around Quantile that accepts a percentage value (0–100) instead
'   of a decimal quantile (0–1).
'
' Parameters:
'   arr : Double()
'       Array of numeric values
'   p : Double
'       Percentile to compute, between 0 and 100 (e.g. 50 for the median)
'   sortedIndices : Variant, Optional
'       Pre-computed sort indices from ArgSort
'
' Returns:
'   Variant
'       The percentile value as a Double, or CVErr(xlErrNum) if p is
'       outside [0, 100]
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
'   Calculates multiple percentiles of a Double array in a single call.
'   Convenience wrapper around Quantiles that accepts percentage values
'   (0–100) rather than decimal quantile levels (0–1).
'
' Parameters:
'   arr : Double()
'       Array of numeric values
'   ps : Double()
'       Array of percentile levels, each between 0 and 100
'   sortedIndices : Variant, Optional
'       Pre-computed sort indices from ArgSort
'
' Returns:
'   Variant()
'       Array of percentile values with the same bounds as ps
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
'   Counts the number of missing or error elements in a Variant array.
'   An element is considered missing if it is an Excel error value, is
'   missing (IsMissing), or converts to an empty string.
'
' Parameters:
'   arr : Variant()
'       Array to inspect for missing values
'
' Returns:
'   Integer
'       The number of missing/error elements
'===============================================================================
Function CountMissing( _
    arr() As Variant _
) As Integer
    ' Counts the number of missing elements (xlErrNA)
    Dim count As Integer
    count = 0
    
    Dim element As Variant
    For Each element In arr
        count = count - CInt(IsError(element) Or IsMissing(element) Or CStr(element) = "")
    Next element
    
    CountMissing = count
End Function

'===============================================================================
' [FUNCTION] CountAboveCritical
'===============================================================================
' Description:
'   Counts the number of elements in a Variant array that are strictly
'   greater than (or optionally greater than or equal to) a critical value.
'   Skips error values. Used by DataRowCls.Describe to count bound
'   violations.
'
' Parameters:
'   arr : Variant()
'       Array of values to test
'   criticalValue : Variant
'       The threshold to compare against. If this is an error value,
'       returns 0 immediately.
'   inclusive : Boolean, Optional
'       If True, counts elements >= criticalValue. If False (default),
'       counts elements > criticalValue. Default: False
'
' Returns:
'   Integer
'       The count of qualifying elements
'===============================================================================
Function CountAboveCritical( _
    arr() As Variant, _
    criticalValue As Variant, _
    Optional inclusive As Boolean = False _
) As Integer
    ' Counts the number of elements of "arr" greater than "criticalValue"
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
'   Counts the number of elements in a Variant array that are strictly less
'   than (or optionally less than or equal to) a critical value. Skips
'   error values. Used by DataRowCls.Describe to count bound violations.
'
' Parameters:
'   arr : Variant()
'       Array of values to test
'   criticalValue : Variant
'       The threshold to compare against. If this is an error value,
'       returns 0 immediately.
'   inclusive : Boolean, Optional
'       If True, counts elements <= criticalValue. If False (default),
'       counts elements < criticalValue. Default: False
'
' Returns:
'   Integer
'       The count of qualifying elements
'===============================================================================
Function CountBelowCritical( _
    arr() As Variant, _
    criticalValue As Variant, _
    Optional inclusive As Boolean = False _
) As Integer
    ' Counts the number of elements of "arr" smaller than "criticalValue"
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
