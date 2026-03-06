'===============================================================================
' [MACRO] Worksheet_Change
'===============================================================================
' Description:
'   Event handler triggered when cells in the Data worksheet are modified.
'   Detects changes to the macro dropdown column and executes the selected
'   macro via MacroUtils.RunMacroFromDropdown.
'
' Parameters:
'   target : Range
'       The cell or range that was changed
'===============================================================================
Private Sub Worksheet_Change( _
    ByVal target As Range _
)
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()
    
    Dim macroColNum As Long
    macroColNum = Me.Range(Specs.MacroColumn & "1").Column
    
    If macroColNum = target.Column Then
        Call MacroUtils.RunMacroFromDropdown(target)
    End If
End Sub