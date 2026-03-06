Option Explicit

' Column where step is located (e.g. "A")
Private pStepColumn As String
' Column where name is located (e.g. "B")
Private pNameColumn As String
' Column where description is located (e.g. "C")
Private pDescColumn As String
' Column where unit is located (e.g. "D")
Private pUnitColumn As String
' Column where key is located (e.g. "E")
Private pKeyColumn As String
' Column where the target is located (e.g. "F")
Private pTargetColumn As String
' Column where min is located (e.g. "G")
Private pMinColumn As String
' Column where max is located (e.g. "H")
Private pMaxColumn As String
' Column where macros are located (e.g. "I")
Private pMacroColumn As String
' Column where data starts (e.g. "J")
Private pDataStartColumn As String

' Row where data starts (e.g. 3)
Private pDataStartRow As Long
' Maximum number of columns to read
Private pNumColumns As Long
' Maximum number of rows to read
Private pNumRows As Long

' Batch number row key (e.g. "[00:00] Batchnummer")
Private pBatchRowKey As String
' Start date row key (e.g. "[00:00] Startdato")
Private pStartDateKey As String

' Sheet name to place plots
Private pPlotsSheetName As String

'===============================================================================
' [SUB] Class_Initialize
'===============================================================================
' Description:
'   Sets all default configuration values for the SpecsCls instance.
'   Called automatically when "New SpecsCls" is used. Defines which columns
'   in the Data worksheet hold step, name, description, unit, key, target,
'   min, max, macro, and data values, as well as limits on rows/columns.
'
' Notes:
'   - Column letters and row numbers reflect the current Data worksheet
'     layout. Update here if the sheet structure changes.
'===============================================================================
Public Sub Class_Initialize()
    pStepColumn = "A"
    pNameColumn = "B"
    pDescColumn = "C"
    pUnitColumn = "D"
    pKeyColumn = "E"
    pTargetColumn = "F"
    pMinColumn = "G"
    pMaxColumn = "H"
    pMacroColumn = "I"
    pDataStartColumn = "J"
    
    pDataStartRow = 3
    pNumColumns = 1000
    pNumRows = 175
    
    pBatchRowKey = "[00:00] Batchnummer"
    pStartDateKey = "[00:00] Startdato"
    
    pPlotsSheetName = "Plots"
End Sub

'===============================================================================
' [PROPERTY] StepColumn
'===============================================================================
' Description:
'   Returns the column letter containing the process step identifier.
'
' Returns:
'   String
'       Column letter, e.g. "A"
'===============================================================================
Public Property Get StepColumn() As String
    StepColumn = pStepColumn
End Property

'===============================================================================
' [PROPERTY] NameColumn
'===============================================================================
' Description:
'   Returns the column letter containing the measurement parameter name.
'
' Returns:
'   String
'       Column letter, e.g. "B"
'===============================================================================
Public Property Get NameColumn() As String
    NameColumn = pNameColumn
End Property

'===============================================================================
' [PROPERTY] DescColumn
'===============================================================================
' Description:
'   Returns the column letter containing the optional row description.
'
' Returns:
'   String
'       Column letter, e.g. "C"
'===============================================================================
Public Property Get DescColumn() As String
    DescColumn = pDescColumn
End Property

'===============================================================================
' [PROPERTY] UnitColumn
'===============================================================================
' Description:
'   Returns the column letter containing the unit of measurement.
'
' Returns:
'   String
'       Column letter, e.g. "D"
'===============================================================================
Public Property Get UnitColumn() As String
    UnitColumn = pUnitColumn
End Property

'===============================================================================
' [PROPERTY] KeyColumn
'===============================================================================
' Description:
'   Returns the column letter where generated key values are written back
'   to the worksheet by DataRowCls.WriteKey.
'
' Returns:
'   String
'       Column letter, e.g. "E"
'===============================================================================
Public Property Get KeyColumn() As String
    KeyColumn = pKeyColumn
End Property

'===============================================================================
' [PROPERTY] TargetColumn
'===============================================================================
' Description:
'   Returns the column letter containing the target (goal) value for each
'   measurement row.
'
' Returns:
'   String
'       Column letter, e.g. "F"
'===============================================================================
Public Property Get TargetColumn() As String
    TargetColumn = pTargetColumn
End Property

'===============================================================================
' [PROPERTY] MinColumn
'===============================================================================
' Description:
'   Returns the column letter containing the minimum acceptable value for
'   each measurement row.
'
' Returns:
'   String
'       Column letter, e.g. "G"
'===============================================================================
Public Property Get MinColumn() As String
    MinColumn = pMinColumn
End Property

'===============================================================================
' [PROPERTY] MaxColumn
'===============================================================================
' Description:
'   Returns the column letter containing the maximum acceptable value for
'   each measurement row.
'
' Returns:
'   String
'       Column letter, e.g. "H"
'===============================================================================
Public Property Get MaxColumn() As String
    MaxColumn = pMaxColumn
End Property

'===============================================================================
' [PROPERTY] MacroColumn
'===============================================================================
' Description:
'   Returns the column letter where the dropdown macro selector is placed
'   for each data row. Users pick an action from this cell's dropdown to
'   trigger sorting, plotting, or descriptive statistics.
'
' Returns:
'   String
'       Column letter, e.g. "I"
'===============================================================================
Public Property Get MacroColumn() As String
    MacroColumn = pMacroColumn
End Property

'===============================================================================
' [PROPERTY] DataStartColumn
'===============================================================================
' Description:
'   Returns the column letter where batch measurement values begin. All
'   columns from here onward (up to NumColumns) are treated as data.
'
' Returns:
'   String
'       Column letter, e.g. "J"
'===============================================================================
Public Property Get DataStartColumn() As String
    DataStartColumn = pDataStartColumn
End Property

'===============================================================================
' [PROPERTY] DataStartRow
'===============================================================================
' Description:
'   Returns the first row number in the Data worksheet that contains
'   measurement parameter rows (i.e. the row after any header rows).
'
' Returns:
'   Long
'       Row number, e.g. 3
'===============================================================================
Public Property Get DataStartRow() As Long
    DataStartRow = pDataStartRow
End Property

'===============================================================================
' [PROPERTY] NumColumns
'===============================================================================
' Description:
'   Returns the maximum number of data columns to read per row. This limits
'   how many batches are loaded from the worksheet.
'
' Returns:
'   Long
'       Maximum column count, e.g. 1000
'===============================================================================
Public Property Get NumColumns() As Long
    NumColumns = pNumColumns
End Property

'===============================================================================
' [PROPERTY] NumRows
'===============================================================================
' Description:
'   Returns the maximum number of rows to scan when loading data from the
'   worksheet. Scanning stops earlier if empty rows are encountered.
'
' Returns:
'   Long
'       Maximum row count, e.g. 175
'===============================================================================
Public Property Get NumRows() As Long
    NumRows = pNumRows
End Property

'===============================================================================
' [PROPERTY] BatchRowKey
'===============================================================================
' Description:
'   Returns the key string used to identify the batch number row in
'   ParsedDataCls. This row is used as the x-axis label when creating
'   scatter plots.
'
' Returns:
'   String
'       Key string, e.g. "[00:00] Batchnummer"
'===============================================================================
Public Property Get BatchRowKey() As String
    BatchRowKey = pBatchRowKey
End Property

'===============================================================================
' [PROPERTY] StartDateKey
'===============================================================================
' Description:
'   Returns the key string used to identify the start date row in
'   ParsedDataCls. This row is used to determine the date interval mask
'   when filtering data for plots.
'
' Returns:
'   String
'       Key string, e.g. "[00:00] Startdato"
'===============================================================================
Public Property Get StartDateKey() As String
    StartDateKey = pStartDateKey
End Property

'===============================================================================
' [PROPERTY] PlotsSheetName
'===============================================================================
' Description:
'   Returns the name of the worksheet where generated charts are placed.
'
' Returns:
'   String
'       Sheet name, e.g. "Plots"
'===============================================================================
Public Property Get PlotsSheetName() As String
    PlotsSheetName = pPlotsSheetName
End Property
