Private Sub Worksheet_Change(ByVal target As Range)
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()
    
    ' Check if the changed cell is in the MacroColumn
    Dim macroColNum As Long
    macroColNum = Me.Range(Specs.MacroColumn & "1").Column
    
    If macroColNum = target.Column Then
        Call MacroUtils.RunMacroFromDropdown(target)
    End If
End Sub
