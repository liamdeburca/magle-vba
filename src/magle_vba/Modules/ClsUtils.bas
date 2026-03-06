' Macros and functions for retrievng the current Data instance

Private gParsedData As ParsedDataCls
Private gSpecs As SpecsCls

'===============================================================================
' [FUNCTION] PauseUpdates
'===============================================================================
' Description:
'   Disables Excel screen updating, event handling, and automatic
'   recalculation. Call this before performing bulk worksheet operations to
'   significantly improve performance. Always pair with ResumeUpdates.
'===============================================================================
Private Function PauseUpdates()
    ' Pauses automatic calculation and screen updates
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Function

'===============================================================================
' [FUNCTION] ResumeUpdates
'===============================================================================
' Description:
'   Re-enables Excel screen updating, event handling, and automatic
'   recalculation. Must be called after PauseUpdates once bulk operations
'   are complete so Excel returns to normal interactive behaviour.
'===============================================================================
Private Function ResumeUpdates()
    ' Resumes automatic calculation and screen updates
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Function

'===============================================================================
' [FUNCTION] GetSpecs
'===============================================================================
' Description:
'   Returns the module-level singleton SpecsCls instance, creating it if
'   it does not yet exist or if a refresh is requested. SpecsCls holds all
'   configuration information about the Data worksheet layout.
'
' Parameters:
'   refresh : Boolean, Optional
'       If True, creates a fresh SpecsCls instance even if one already
'       exists. Default: False
'
' Returns:
'   SpecsCls
'       The current (or newly created) SpecsCls singleton
'===============================================================================
Function GetSpecs(Optional refresh As Boolean = False) As SpecsCls

    If gSpecs Is Nothing Or refresh Then
        Set gSpecs = New SpecsCls
    End If
    
    Set GetSpecs = gSpecs
End Function

'===============================================================================
' [FUNCTION] GetParsedData
'===============================================================================
' Description:
'   Returns the module-level singleton ParsedDataCls instance, creating and
'   initialising it from the Data worksheet if it does not yet exist or if
'   a refresh is requested. Screen updates and events are paused during
'   initialisation for performance.
'
' Parameters:
'   refresh : Boolean, Optional
'       If True, re-reads the worksheet and creates a fresh ParsedDataCls
'       instance. Default: False
'
' Returns:
'   ParsedDataCls
'       The current (or newly created) ParsedDataCls singleton
'===============================================================================
Function GetParsedData(Optional refresh As Boolean = False) As ParsedDataCls
    
    If gParsedData Is Nothing Or refresh Then
        ' Instantiate new ParsedDataCls class
        Call PauseUpdates
        
        Set gParsedData = New ParsedDataCls
        Call gParsedData.Init
        
        Call ResumeUpdates
    End If
    
    Set GetParsedData = gParsedData
End Function

' ----- MACROS ------ '

'===============================================================================
' [MACRO] RefreshSpecs
'===============================================================================
' Description:
'   Resets and re-initialises the global SpecsCls singleton. Run this if
'   the Data worksheet layout has changed and the cached configuration needs
'   to be updated.
'===============================================================================
Sub RefreshSpecs()
    ' Resets and re-initialises the current SpecsCls instance
    Call GetSpecs(True)
End Sub

'===============================================================================
' [MACRO] RefreshParsedData
'===============================================================================
' Description:
'   Resets and re-initialises the global ParsedDataCls singleton by
'   re-reading the Data worksheet. Use this after manually editing data
'   in the sheet to pick up changes without reopening the workbook.
'===============================================================================
Sub RefreshParsedData()
    ' Resets and re-initialises the current Data instance
    Call GetParsedData(True)
End Sub

'===============================================================================
' [FUNCTION] IsValidStepFormat
'===============================================================================
' Description:
'   Validates that a step string conforms to the required "xx:xx" format,
'   where "x" is any character and each side of the colon is exactly two
'   characters long. Used by DataRowCls.LoadFromSheet to validate step
'   values before loading a row.
'
' Parameters:
'   stepValue : String
'       The step string to validate
'
' Returns:
'   Boolean
'       True if the string matches "xx:xx" format, False otherwise
'===============================================================================
Function IsValidStepFormat(stepValue As String) As Boolean
    ' Check if step matches format xx:xx where x is any character
    ' Using pattern: *:* with exactly 2 chars before and after colon
    Dim pattern As String
    pattern = stepValue
    
    ' Should have exactly one colon
    If Len(pattern) - Len(Replace(pattern, ":", "")) <> 1 Then
        IsValidStepFormat = False
        Exit Function
    End If
    
    ' Check that we have exactly 2 chars, colon, 2 chars
    Dim parts() As String
    parts = Split(pattern, ":")
    
    If UBound(parts) <> 1 Then
        IsValidStepFormat = False
        Exit Function
    End If
    
    If Len(parts(0)) <> 2 Or Len(parts(1)) <> 2 Then
        IsValidStepFormat = False
        Exit Function
    End If
    
    IsValidStepFormat = True
End Function

'===============================================================================
' [SUB] ClearKeys
'===============================================================================
' Description:
'   Removes all generated key values from the key column of the Data
'   worksheet. Only clears cells in rows where both step and name are
'   present (i.e. valid data rows). Optionally pauses screen updates for
'   performance.
'
' Parameters:
'   pause : Boolean, Optional
'       If True (default), pauses Excel updates before clearing and resumes
'       afterwards
'===============================================================================
Public Sub ClearKeys(Optional pause As Boolean = True)
    ' Clears the key-column of all generated values
    If pause Then Call PauseUpdates
    
    Dim Specs As SpecsCls
    Set Specs = ClsUtils.GetSpecs()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")
    
    Dim N As Long, M As Long
    N = Specs.DataStartRow
    M = N + Specs.NumRows - 1
    
    Dim stepCell As Range, nameCell As Range, keyCell As Range
    Dim i As Long
    
    For i = N To M
        stepValue = ws.Range(Specs.StepColumn & i).value
        nameValue = ws.Range(Specs.NameColumn & i).value
        Set keyCell = ws.Range(Specs.KeyColumn & i)
        
        If Not (IsEmpty(stepValue) Or stepValue = "") And Not (IsEmpty(nameValue) Or nameValue = "") Then
            ' key-value possible -> clear it
            Call keyCell.ClearContents
        End If
    Next i
    
    If pause Then Call ResumeUpdates
End Sub

'===============================================================================
' [SUB] ClearConditionalFormatting
'===============================================================================
' Description:
'   Removes all conditional formatting rules from the data range of every
'   valid row in the Data worksheet. Optionally pauses screen updates for
'   performance.
'
' Parameters:
'   pause : Boolean, Optional
'       If True (default), pauses Excel updates before clearing and resumes
'       afterwards
'===============================================================================
Public Sub ClearConditionalFormatting(Optional pause As Boolean = True)
    'Clears conditional formatting
    If pause Then Call PauseUpdates
    
    Dim Specs As SpecsCls
    Set Specs = ClsUtils.GetSpecs()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")

    Dim dataStartCol As Long, dataEndCol As Long
    dataStartCol = ws.Range(Specs.DataStartColumn & "1").Column
    dataEndCol = dataStartCol + Specs.NumColumns - 1
    
    Dim dataRange As Range
    
    Dim N As Long, M As Long
    N = Specs.DataStartRow
    M = N + Specs.NumRows - 1
    
    
    Dim stepCell As Range, nameCell As Range, keyCell As Range
    Dim i As Long
    
    For i = N To M
        stepValue = ws.Range(Specs.StepColumn & i).value
        nameValue = ws.Range(Specs.NameColumn & i).value
        Set keyCell = ws.Range(Specs.KeyColumn & i)
        
        If Not (IsEmpty(stepValue) Or stepValue = "") And Not (IsEmpty(nameValue) Or nameValue = "") Then
            ' key-value possible -> clear all formatting rules
            Set dataRange = ws.Range(ws.Cells(i, dataStartCol), ws.Cells(i, dataEndCol))
            Call dataRange.FormatConditions.Delete
        End If
    Next i
    
    If pause Then Call ResumeUpdates
End Sub

'===============================================================================
' [SUB] ClearMacros
'===============================================================================
' Description:
'   Removes the dropdown macro selector and clears the macro cell content
'   from every valid row in the Data worksheet. Optionally pauses screen
'   updates for performance.
'
' Parameters:
'   pause : Boolean, Optional
'       If True (default), pauses Excel updates before clearing and resumes
'       afterwards
'===============================================================================
Public Sub ClearMacros(Optional pause As Boolean = True)
    ' Clears macro-dropdown menus
    If pause Then Call PauseUpdates
    
    Dim Specs As SpecsCls
    Set Specs = ClsUtils.GetSpecs()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")
    
    Dim N As Long, M As Long
    N = Specs.DataStartRow
    M = N + Specs.NumRows - 1
    
    Dim stepCell As Range, nameCell As Range, keyCell As Range
    Dim macroCell As Range
    Dim i As Long
    
    For i = N To M
        stepValue = ws.Range(Specs.StepColumn & i).value
        nameValue = ws.Range(Specs.NameColumn & i).value
        Set keyCell = ws.Range(Specs.KeyColumn & i)
        
        If Not (IsEmpty(stepValue) Or stepValue = "") And Not (IsEmpty(nameValue) Or nameValue = "") Then
            ' key-value possible -> clear it
            Set macroCell = ws.Range(Specs.MacroColumn & i)
            
            With macroCell
                .ClearContents
                .Validation.Delete
            End With
        End If
    Next i
    
    If pause Then Call ResumeUpdates
End Sub

'===============================================================================
' [MACRO] Reset
'===============================================================================
' Description:
'   Resets the Data worksheet to a clean state by running ClearKeys,
'   ClearConditionalFormatting, and ClearMacros in sequence. Screen updates
'   are paused for the duration of the operation.
'
'   Typically called on workbook close (Workbook_BeforeClose) or via the
'   README sheet interface.
'===============================================================================
Public Sub Reset()
    ' Runs the following macros:
    ' 1. ClearKeys
    ' 2. ClearConditionalFormatting
    ' 3. ClearMacros
    Call PauseUpdates
    
    Call ClearKeys(False)
    Call ClearConditionalFormatting(False)
    Call ClearMacros(False)
    
    Call ResumeUpdates
End Sub

'===============================================================================
' [MACRO] Start
'===============================================================================
' Description:
'   Loads data from the Data worksheet into a new ParsedDataCls instance.
'   Writes keys, applies conditional formatting, and adds macro dropdowns
'   to all valid rows. Screen updates are paused for performance.
'
'   Typically called on workbook open (Workbook_Open) or via the README
'   sheet interface.
'===============================================================================
Public Sub Start()
    ' Instantiates a ParsedDataCls class
    Call PauseUpdates
    
    Set gParsedData = New ParsedDataCls
    Call gParsedData.Init

    Call ResumeUpdates
End Sub

'===============================================================================
' [MACRO] ResetAndStart
'===============================================================================
' Description:
'   Convenience macro that runs Reset followed by Start. Completely clears
'   and re-initialises the Data worksheet — useful when the sheet contents
'   have changed significantly and a full reload is needed.
'===============================================================================
Public Sub ResetAndStart()
    ' Resets and re-instantiates
    Call Reset
    Call Start
End Sub
