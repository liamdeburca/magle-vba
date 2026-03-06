' File containing plotting macros

'===============================================================================
' [FUNCTION] GetRowsForCell
'===============================================================================
' Description:
'   Returns a Collection of the three DataRowCls instances needed to create
'   a standard scatter plot for the specified cell: the batch-number row
'   (x-axis labels), the measurement row for the cell's worksheet row
'   (y-axis values), and the start-date row (secondary x-axis labels).
'
' Parameters:
'   cell : Range
'       Any cell in the Data worksheet; its row number identifies the
'       measurement DataRowCls to include
'
' Returns:
'   Collection
'       A 3-item Collection: (1) batch row, (2) measurement row,
'       (3) start-date row
'===============================================================================
Private Function GetRowsForCell(cell As Range) As Collection
    Dim Specs As SpecsCls
    Set Specs = GetSpecs()
    
    Dim ParsedData As ParsedDataCls
    Set ParsedData = GetParsedData()
    
    Dim Rows As New Collection
    Rows.Add ParsedData.GetRowFromKey(Specs.BatchRowKey)
    Rows.Add ParsedData.GetRowFromIndex(cell.Row)
    Rows.Add ParsedData.GetRowFromKey(Specs.StartDateKey)
    
    Set GetRowsForCell = Rows
End Function

'===============================================================================
' [FUNCTION] IsValidStat
'===============================================================================
' Description:
'   Validates that a statistic string is one of the accepted values:
'   empty string (no statistic), "mean", or "median". Used to sanitise
'   user input in BasicScatterPlot_.
'
' Parameters:
'   s : String
'       The statistic string to check (case- and whitespace-insensitive)
'
' Returns:
'   Boolean
'       True if the string is a valid statistic option, False otherwise
'===============================================================================
Private Function IsValidStat(s As String) As Boolean
    s = Trim(LCase(s))
    IsValidStat = (s = "") Or (s = "mean") Or (s = "median")
End Function

' ===== SCATTER PLOT FOR CURRENT ROW =====
'===============================================================================
' [SUB] BasicScatterPlot_
'===============================================================================
' Description:
'   Creates a basic scatter plot for the measurement row corresponding to
'   the given cell. The user is optionally prompted to select a year filter
'   (via DataRowCls.GetDateIntervalMask) and a statistic line (mean or
'   median). The chart is placed on the Plots worksheet.
'
' Parameters:
'   cell : Range
'       A cell in the Data worksheet whose row determines which
'       measurement parameter to plot
'
' Notes:
'   - Called by MacroUtils.RunMacroFromDropdown when "Scatter" is selected
'   - Delegates chart creation to BasicScatterPlot.Run
'===============================================================================
Public Sub BasicScatterPlot_(cell As Range)
    ' Creates a basic scatter plot for the given row.
    ' (Optional) Asks the user if they want to plot the data for a specific year.
    Dim Specs As SpecsCls
    Set Specs = ClsUtils.GetSpecs()
    
    Dim Rows As Collection
    Set Rows = GetRowsForCell(cell)
    
    Dim xDR As DataRowCls, yDR As DataRowCls, x2DR As DataRowCls
    Set xDR = Rows(1)
    Set yDR = Rows(2)
    Set x2DR = Rows(3)
    
    Dim mask() As Boolean
    mask = x2DR.GetDateIntervalMask()
        
    '' Statistic
    Dim stat As String
    
    userInput = InputBox("Enter which statistic to plot (<Enter> none, <mean>, or <median>)")
    stat = Trim(LCase(userInput))
    If Not IsValidStat(stat) Then stat = ""
    
    Call BasicScatterPlot.Run( _
        xDR, _
        yDR, _
        x2DataRow:=x2DR, _
        mask:=mask, _
        stat:=stat, _
        targetSheet:=Specs.PlotsSheetName _
    )
End Sub


