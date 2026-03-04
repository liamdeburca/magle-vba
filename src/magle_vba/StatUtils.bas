Option Explicit

' Statistical functions for working with double arrays containing potential NA errors

' ============================================================================
' UTILITY FUNCTIONS
' ============================================================================

'''
' Removes NA errors from an array (variant), returning a filtered array (double)
' @param arr: Array potentially containing #N/A errors
' @return: New array with NA values removed
'
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

'''
' Returns array of indices that would sort the input array
' Uses quicksort algorithm
' @param arr: Array to sort
' @return: Array of indices (0-based) representing sorted order
'
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

'''
' Sorts an array using provided or calculated sort indices
' @param arr: Array to sort
' @param sortedIndices: Optional array of indices. If not provided, calculated via ArrayArgSort
' @return: Sorted array
'
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

'''
' Calculates the arithmetic mean of an array
' @param arr: Array of doubles
' @return: Mean value
'
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

'''
' Calculates the standard deviation of an array
' @param arr: Array of doubles
' @param ddof: Degrees of freedom correction (default 1 for sample std dev)
' @return: Standard deviation
'
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

'''
' Calculates a single quantile of an array
' Uses linear interpolation (Type 7 - R default)
' @param arr: Array of doubles
' @param p: Quantile between 0 and 1
' @return: Quantile value
'
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

'''
' Calculates multiple quantiles of an array
' @param arr: Array of doubles
' @param ps: Array of quantiles between 0 and 1
' @return: Array of quantile values
'
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

'''
' Calculates a single percentile of an array
' Wrapper around ArrayQuantile (converts percentage to decimal)
' @param arr: Array of doubles
' @param p: Percentile between 0 and 100
' @return: Percentile value
'
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

'''
' Calculates multiple percentiles of an array
' Wrapper around ArrayQuantiles (converts percentages to decimals)
' @param arr: Array of doubles
' @param ps: Array of percentiles between 0 and 100
' @return: Array of percentile values
'
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
