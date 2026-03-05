'===============================================================================
' [MACRO] Worksheet_Change
'===============================================================================
' Description:
'   Event handler for the README worksheet. Provides a simple interface for
'   running top-level macros by entering "Run" in column B next to a macro
'   name in column A. Supports Reset, Start, and ClearPlots commands.
'
' Parameters:
'   target : Range
'       The cell that was changed
'===============================================================================
Private Sub Worksheet_Change( _
    ByVal target As Range _
)
    If target.Column = 2 And CStr(target) = "Run" Then
        Dim macroName As String
        macroName = CStr(target.Offset(0, -1))
        
        Select Case macroName
            Case "Reset"
                ' Call ClsUtils.Reset
            
            Case "Start"
                ' Call ClsUtils.Start
                
            Case "ClearPlots"
                ' Call PlottingUtils.ClearPlots
                
            Case Else
        End Select
        
        target.ClearContents
    End If
End Sub
