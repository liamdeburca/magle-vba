'===============================================================================
' Module: ClsUtils
'===============================================================================
' Description:
'   Core utility module providing singleton access to SpecsCls and
'   ParsedDataCls instances, along with worksheet management macros. This
'   module serves as the central entry point for initializing and resetting
'   the data processing system.
'
'   Key functions:
'   - GetSpecs: Returns the singleton SpecsCls configuration instance
'   - GetParsedData: Returns the singleton ParsedDataCls data instance
'   - Reset: Clears all generated content from the Data worksheet
'   - Start: Initializes data loading and processing
'===============================================================================

Private gParsedData As ParsedDataCls
Private gSpecs As SpecsCls

'===============================================================================
' [SUB] PauseUpdates
'===============================================================================
' Description:
'   Disables screen updating, events, and automatic calculation for
'   performance during batch operations.
'===============================================================================
Private Sub PauseUpdates()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

'===============================================================================
' [SUB] ResumeUpdates
'===============================================================================
' Description:
'   Re-enables screen updating, events, and automatic calculation after
'   batch operations complete.
'===============================================================================
Private Sub ResumeUpdates()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Sub

'===============================================================================
' [FUNCTION] GetSpecs
'===============================================================================
' Description:
'   Returns the singleton SpecsCls instance containing worksheet layout
'   configuration. Creates a new instance on first call or when refresh
'   is requested.
'
' Parameters:
'   refresh : Boolean, Optional
'       If True, creates a new instance instead of returning cached one.
'       Default: False
'
' Returns:
'   SpecsCls
'       The configuration instance
'===============================================================================
Function GetSpecs( _
    Optional refresh As Boolean = False _
) As SpecsCls
    If gSpecs Is Nothing Or refresh Then
        Set gSpecs = New SpecsCls
    End If
    
    Set GetSpecs = gSpecs
End Function

'===============================================================================
' [FUNCTION] GetParsedData
'===============================================================================
' Description:
'   Returns the singleton ParsedDataCls instance containing all loaded data
'   rows. Creates and initializes a new instance on first call or when
'   refresh is requested.
'
' Parameters:
'   refresh : Boolean, Optional
'       If True, reloads data from the worksheet. Default: False
'
' Returns:
'   ParsedDataCls
'       The data container instance
'===============================================================================
Function GetParsedData( _
    Optional refresh As Boolean = False _
) As ParsedDataCls
    If gParsedData Is Nothing Or refresh Then
        Call PauseUpdates
        
        Set gParsedData = New ParsedDataCls
        Call gParsedData.Init
        
        Call ResumeUpdates
    End If
    
    Set GetParsedData = gParsedData
End Function

'===============================================================================
' [MACRO] RefreshSpecs
'===============================================================================
' Description:
'   Forces recreation of the SpecsCls singleton instance.
'===============================================================================
Sub RefreshSpecs()
    Call GetSpecs(True)
End Sub

'===============================================================================
' [MACRO] RefreshParsedData
'===============================================================================
' Description:
'   Forces reload of all data from the Data worksheet.
'===============================================================================
Sub RefreshParsedData()
    Call GetParsedData(True)
End Sub

'===============================================================================
' [FUNCTION] IsValidStepFormat
'===============================================================================
' Description:
'   Validates that a step value matches the required format "xx:xx" where
'   each x is any character.
'
' Parameters:
'   stepValue : String
'       The step value to validate
'
' Returns:
'   Boolean
'       True if format is valid, False otherwise
'===============================================================================
Function IsValidStepFormat( _
    stepValue As String _
) As Boolean
    Dim pattern As String
    pattern = stepValue
    
    If Len(pattern) - Len(Replace(pattern, ":", "")) <> 1 Then
        IsValidStepFormat = False
        Exit Function
    End If
    
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
'   Removes all generated key values from the Key column in the Data
'   worksheet. Keys are regenerated when data is loaded via GetParsedData.
'
' Parameters:
'   pause : Boolean, Optional
'       If True, pauses screen updates during operation. Default: True
'===============================================================================
Public Sub ClearKeys( _
    Optional pause As Boolean = True _
)
    If pause Then Call PauseUpdates
    
    Dim Specs As SpecsCls
    Set Specs = ClsUtils.GetSpecs()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")
    
    Dim N As Long
    Dim M As Long
    N = Specs.DataStartRow
    M = N + Specs.NumRows - 1
    
    Dim i As Long
    For i = N To M
        Dim stepValue As Variant
        Dim nameValue As Variant
        stepValue = ws.Range(Specs.StepColumn & i).value
        nameValue = ws.Range(Specs.NameColumn & i).value
        
        Dim keyCell As Range
        Set keyCell = ws.Range(Specs.KeyColumn & i)
        
        If Not (IsEmpty(stepValue) Or stepValue = "") And _
           Not (IsEmpty(nameValue) Or nameValue = "") Then
            Call keyCell.ClearContents
        End If
    Next i
    
    If pause Then Call ResumeUpdates
End Sub

'===============================================================================
' [SUB] ClearConditionalFormatting
'===============================================================================
' Description:
'   Removes all conditional formatting rules from data cells in the Data
'   worksheet. Formatting is reapplied when data is loaded via GetParsedData.
'
' Parameters:
'   pause : Boolean, Optional
'       If True, pauses screen updates during operation. Default: True
'===============================================================================
Public Sub ClearConditionalFormatting( _
    Optional pause As Boolean = True _
)
    If pause Then Call PauseUpdates
    
    Dim Specs As SpecsCls
    Set Specs = ClsUtils.GetSpecs()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")

    Dim dataStartCol As Long
    Dim dataEndCol As Long
    dataStartCol = ws.Range(Specs.DataStartColumn & "1").Column
    dataEndCol = dataStartCol + Specs.NumColumns - 1
    
    Dim N As Long
    Dim M As Long
    N = Specs.DataStartRow
    M = N + Specs.NumRows - 1
    
    Dim i As Long
    For i = N To M
        Dim stepValue As Variant
        Dim nameValue As Variant
        stepValue = ws.Range(Specs.StepColumn & i).value
        nameValue = ws.Range(Specs.NameColumn & i).value
        
        If Not (IsEmpty(stepValue) Or stepValue = "") And _
           Not (IsEmpty(nameValue) Or nameValue = "") Then
            Dim dataRange As Range
            Set dataRange = ws.Range( _
                ws.Cells(i, dataStartCol), _
                ws.Cells(i, dataEndCol) _
            )
            Call dataRange.FormatConditions.Delete
        End If
    Next i
    
    If pause Then Call ResumeUpdates
End Sub

'===============================================================================
' [SUB] ClearMacros
'===============================================================================
' Description:
'   Removes all macro dropdown menus from the Macro column in the Data
'   worksheet. Menus are recreated when data is loaded via GetParsedData.
'
' Parameters:
'   pause : Boolean, Optional
'       If True, pauses screen updates during operation. Default: True
'===============================================================================
Public Sub ClearMacros( _
    Optional pause As Boolean = True _
)
    If pause Then Call PauseUpdates
    
    Dim Specs As SpecsCls
    Set Specs = ClsUtils.GetSpecs()
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Data")
    
    Dim N As Long
    Dim M As Long
    N = Specs.DataStartRow
    M = N + Specs.NumRows - 1
    
    Dim i As Long
    For i = N To M
        Dim stepValue As Variant
        Dim nameValue As Variant
        stepValue = ws.Range(Specs.StepColumn & i).value
        nameValue = ws.Range(Specs.NameColumn & i).value
        
        If Not (IsEmpty(stepValue) Or stepValue = "") And _
           Not (IsEmpty(nameValue) Or nameValue = "") Then
            Dim macroCell As Range
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
'   Clears all generated content from the Data worksheet including keys,
'   conditional formatting, and macro dropdowns. Returns the worksheet to
'   a clean state ready for re-initialization.
'===============================================================================
Public Sub Reset()
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
'   Initializes the data processing system by creating a new ParsedDataCls
'   instance and loading all data from the Data worksheet.
'===============================================================================
Public Sub Start()
    Set gParsedData = New ParsedDataCls
    Call gParsedData.Init
End Sub

'===============================================================================
' [MACRO] ResetAndStart
'===============================================================================
' Description:
'   Performs a complete system reset followed by initialization. Clears all
'   generated content and then reloads data from the worksheet.
'===============================================================================
Public Sub ResetAndStart()
    Call Reset
    Call Start
End Sub
