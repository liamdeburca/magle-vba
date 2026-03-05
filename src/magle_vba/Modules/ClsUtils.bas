' Macros and functions for retrievng the current Data instance

Private gParsedData As ParsedDataCls
Private gSpecs As SpecsCls

Private Function PauseUpdates()
    ' Pauses automatic calculation and screen updates
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Function

Private Function ResumeUpdates()
    ' Resumes automatic calculation and screen updates
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
End Function

Function GetSpecs(Optional refresh As Boolean = False) As SpecsCls

    If gSpecs Is Nothing Or refresh Then
        Set gSpecs = New SpecsCls
    End If
    
    Set GetSpecs = gSpecs
End Function

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

Sub RefreshSpecs()
    ' Resets and re-initialises the current SpecsCls instance
    Call GetSpecs(True)
End Sub

Sub RefreshParsedData()
    ' Resets and re-initialises the current Data instance
    Call GetParsedData(True)
End Sub

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

Public Sub Start()
    ' Instantiates a ParsedDataCls class
    Set gParsedData = New ParsedDataCls
    Call gParsedData.Init
End Sub

Public Sub ResetAndStart()
    ' Resets and re-instantiates
    Call Reset
    Call Start
End Sub
