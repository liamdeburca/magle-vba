'===============================================================================
' [MACRO] Workbook_Open
'===============================================================================
' Description:
'   Event handler that runs automatically when the workbook is opened.
'   Prompts the user to confirm whether they want to load data from the
'   Data worksheet. If confirmed (or Enter is pressed), calls ClsUtils.Start
'   to initialise the ParsedDataCls singleton, write keys, apply conditional
'   formatting, and add macro dropdowns.
'
' Notes:
'   - Pressing Enter or typing "y" (case-insensitive) triggers loading
'   - Any other input skips loading; data can be loaded later via Start
'===============================================================================
Private Sub Workbook_Open()
    Dim userInput As String
    
    userInput = InputBox("Load Data worksheet? [Y/n]")
    If LCase(Trim(userInput)) = "y" Or userInput = "" Then
        Call ClsUtils.Start
    End If
End Sub

'===============================================================================
' [MACRO] Workbook_BeforeClose
'===============================================================================
' Description:
'   Event handler that runs automatically just before the workbook closes.
'   Prompts the user separately to:
'     1. Reset the Data worksheet (clear keys, formatting, dropdowns)
'     2. Clear all charts from the Plots worksheet
'   Each prompt defaults to Yes if the user presses Enter.
'
' Parameters:
'   cancel : Boolean
'       Standard Excel cancel flag for the BeforeClose event (not used)
'
' Notes:
'   - Resetting ensures the workbook is in a clean state when reopened on
'     another machine or by another user
'===============================================================================
Private Sub Workbook_BeforeClose(cancel As Boolean)
    Dim Specs As SpecsCls
    Set Specs = ClsUtils.GetSpecs()
    
    Dim userInput As String
    
    userInput = InputBox("Reset Data worksheet? [Y/n]")
    If LCase(Trim(userInput)) = "y" Or userInput = "" Then
        Call ClsUtils.Reset
    End If

    userInput = InputBox("Clear " & Specs.PlotsSheetName & " worksheet? [Y/n]")
    If LCase(Trim(userInput)) = "y" Or userInput = "" Then
        Call PlottingUtils.ClearPlots
    End If
End Sub

