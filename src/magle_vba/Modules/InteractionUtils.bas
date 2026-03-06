'===============================================================================
' [FUNCTION] IsValidSingleYear
'===============================================================================
' Description:
'   Validates that a string represents a single four-digit year, e.g.
'   "2024". Used internally by GetDateBounds to parse user year input.
'
' Parameters:
'   s : String
'       The string to validate
'
' Returns:
'   Boolean
'       True if the string is exactly four characters and fully numeric
'===============================================================================
Private Function IsValidSingleYear(s As String) As Boolean
    IsValidSingleYear = (Len(s) = 4) And IsNumeric(s)
End Function

'===============================================================================
' [FUNCTION] IsValidRangeOfYears
'===============================================================================
' Description:
'   Validates that a string represents a range of years in "YYYY:YYYY"
'   format where the second year is greater than or equal to the first,
'   e.g. "2020:2024". Used internally by GetDateBounds to parse user input.
'
' Parameters:
'   s : String
'       The string to validate
'
' Returns:
'   Boolean
'       True if the string is a valid year range, False otherwise
'===============================================================================
Private Function IsValidRangeOfYears(s As String) As Boolean
    Dim splits() As String
    
    If Len(s) = 9 And InStr(1, s, ":") Then
        splits = Split(s, ":")
        If IsValidSingleYear(splits(0)) And IsValidSingleYear(splits(1)) Then
            IsValidRangeOfYears = CInt(splits(1)) >= CInt(splits(0))
            Exit Function
        End If
    End If
    
    IsValidRangeOfYears = False
End Function

'===============================================================================
' [FUNCTION] GetDateBounds
'===============================================================================
' Description:
'   Prompts the user to enter a year filter for scatter plots and returns
'   the lower and upper date bounds as an array of two Excel date serial
'   numbers. Three input options are accepted:
'     - Empty / Enter  : include all dates (1900-01-01 to 3000-01-01)
'     - YYYY           : include only the specified calendar year
'     - YYYY:YYYY      : include the inclusive range of calendar years
'   If the input is invalid, the user is notified and all dates are
'   included.
'
' Returns:
'   Double(1 To 2)
'       result(1) is the inclusive lower date bound (as a serial number)
'       result(2) is the exclusive upper date bound (as a serial number)
'
' Example:
'   Dim bounds() As Double
'   bounds = GetDateBounds()
'===============================================================================
Function GetDateBounds() As Double()
    Dim userInput As String
    Dim D1 As Double, D2 As Double

    D1 = CDbl(DateValue("1900-01-01"))
    D2 = CDbl(DateValue("3000-01-01"))
    
    userInput = InputBox("Enter which year(s) to plot (<Enter> all, <YYYY> specific year, or <YYYY:YYYY> range of years)")
    If userInput = "" Then
        ' Include all years
        
    ElseIf IsValidSingleYear(userInput) Then
        ' Include specific year
        D1 = CDbl(DateValue("" & CInt(userInput) & "-01-01"))
        D2 = CDbl(DateValue("" & (CInt(userInput) + 1) & "-01-01"))
        
    ElseIf IsValidRangeOfYears(userInput) Then
        ' Include range of years
        Dim splits() As String
        splits = Split(userInput, ":")
        D1 = CDbl(DateValue(splits(0) & "-01-01"))
        D2 = CDbl(DateValue(splits(1) & "-01-01"))
        
    Else
        Call MsgBox("Invalid selection!")
    End If

    Dim result(1 To 2) As Double
    result(1) = D1
    result(2) = D2

    GetDateBounds = result
End Function

