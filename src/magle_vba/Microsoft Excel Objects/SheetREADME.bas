Private Sub Worksheet_Change(ByVal target As Range)

    If target.Column = 2 And CStr(target) = "Run" Then
        Dim macroName As String
        macroName = CStr(target.Offset(0, -1))
        
        Select Case macroName
            Case "Reset"
                ' Run the Reset macro
                ' Call ClsUtils.Reset
            
            Case "Start"
                ' Run the Start macro
                ' Call ClsUtils.Start
                
            Case "ClearPlots"
                ' Run the ClearPlots macro
                ' Call PlottingUtils.ClearPlots
                
            Case Else
                ' Didnt recognise the macro name
        End Select
        
        ' Clear cell
        target.ClearContents
    End If
End Sub
