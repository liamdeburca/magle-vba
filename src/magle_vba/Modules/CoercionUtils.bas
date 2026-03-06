'===============================================================================
' [FUNCTION] CoerceDouble
'===============================================================================
' Description:
'   Attempts to convert a Variant value to a Double. If conversion fails
'   (e.g. the value is text, empty, or an existing error), returns
'   xlErrValue instead of raising a runtime error. Used throughout
'   DataRowCls when populating numeric data arrays.
'
' Parameters:
'   value : Variant
'       The value to convert
'
' Returns:
'   Variant
'       A Double if conversion succeeds, or CVErr(xlErrValue) otherwise
'===============================================================================
Function CoerceDouble( _
    value As Variant _
) As Variant
    On Error GoTo CoercionError
    CoerceDouble = CDbl(value)
    Exit Function
CoercionError:
    CoerceDouble = CVErr(xlErrValue)
End Function

'===============================================================================
' [FUNCTION] CoerceTimeValue
'===============================================================================
' Description:
'   Attempts to convert a time string (e.g. "14:30") to an Excel time
'   serial number using VBA's TimeValue function. Returns xlErrValue if
'   the string cannot be parsed as a time.
'
' Parameters:
'   value : String
'       A time string to parse, e.g. "14:30"
'
' Returns:
'   Variant
'       A Date (time serial number) if conversion succeeds, or
'       CVErr(xlErrValue) otherwise
'===============================================================================
Function CoerceTimeValue( _
    value As String _
) As Variant
    On Error GoTo CoercionError
    CoerceTimeValue = TimeValue(value)
    Exit Function
CoercionError:
    CoerceTimeValue = CVErr(xlErrValue)
End Function

'===============================================================================
' [FUNCTION] CoerceDateValue
'===============================================================================
' Description:
'   Attempts to convert a date string to an Excel date serial number using
'   VBA's DateValue function. Returns xlErrValue if the string cannot be
'   parsed as a date.
'
' Parameters:
'   value : String
'       A date string to parse, e.g. "2024-03-15"
'
' Returns:
'   Variant
'       A Date (date serial number) if conversion succeeds, or
'       CVErr(xlErrValue) otherwise
'===============================================================================
Function CoerceDateValue( _
    value As String _
) As Variant
    On Error GoTo CoercionError
    CoerceDateValue = DateValue(value)
    Exit Function
CoercionError:
    CoerceDateValue = CVErr(xlErrValue)
End Function

'===============================================================================
' [FUNCTION] CoerceHHMM
'===============================================================================
' Description:
'   Converts an "hh:mm" time string to a Double representing the Excel
'   time serial number. This is used for data rows whose unit is "hh:mm"
'   so that numeric operations (statistics, plotting) can be performed.
'   Returns xlErrValue if the string is not a valid time.
'
' Parameters:
'   value : String
'       A time string in "hh:mm" format, e.g. "08:30"
'
' Returns:
'   Variant
'       A Double time serial (fraction of a day) if successful, or
'       CVErr(xlErrValue) otherwise
'===============================================================================
Function CoerceHHMM( _
    value As String _
) As Variant
    Dim timeVal As Variant
    timeVal = CoerceTimeValue(value)
    If Not IsError(timeVal) Then
        CoerceHHMM = CDbl(timeVal)
    Else
        CoerceHHMM = CVErr(xlErrValue)
    End If
End Function

'===============================================================================
' [FUNCTION] CoerceYYYYMMDD
'===============================================================================
' Description:
'   Converts a date string in "YYYY-MM-DD" (or similar date) format to a
'   Double representing the Excel date serial number. Used for data rows
'   whose unit is "åååå-mm-dd" or "åååå-mm". Returns xlErrValue if the
'   string is not a valid date.
'
' Parameters:
'   value : String
'       A date string, e.g. "2024-03-15"
'
' Returns:
'   Variant
'       A Double date serial number if successful, or CVErr(xlErrValue)
'       otherwise
'===============================================================================
Function CoerceYYYYMMDD( _
    value As String _
) As Variant
    Dim dateVal As Variant
    dateVal = CoerceDateValue(value)
    If Not IsError(dateVal) Then
        CoerceYYYYMMDD = CDbl(dateVal)
    Else
        CoerceYYYYMMDD = CVErr(xlErrValue)
    End If
End Function

