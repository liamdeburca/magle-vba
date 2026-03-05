'===============================================================================
' [CLASS] SpecsCls
'===============================================================================
' Description:
'   Configuration class that defines the structure of the Data worksheet.
'   Contains column letter assignments and row/column counts that tell
'   DataRowCls and ParsedDataCls where to find data elements. This class
'   centralizes worksheet layout configuration to make the system adaptable
'   to different spreadsheet formats.
'
'   Typically accessed via the GetSpecs() function in ClsUtils, which
'   maintains a singleton instance.
'
' Properties:
'   StepColumn : String - Column letter for process step (default "A")
'   NameColumn : String - Column letter for parameter name (default "B")
'   DescColumn : String - Column letter for description (default "C")
'   UnitColumn : String - Column letter for unit (default "D")
'   KeyColumn : String - Column letter for generated key (default "E")
'   TargetColumn : String - Column letter for target value (default "F")
'   MinColumn : String - Column letter for minimum value (default "G")
'   MaxColumn : String - Column letter for maximum value (default "H")
'   MacroColumn : String - Column letter for macro dropdown (default "I")
'   DataStartColumn : String - First column of measurement data (default "J")
'   DataStartRow : Long - First row of data (default 3)
'   NumColumns : Long - Number of data columns to read (default 5)
'   NumRows : Long - Maximum rows to scan (default 75)
'   BatchRowKey : String - Key for batch number row
'   StartDateKey : String - Key for start date row
'   PlotsSheetName : String - Name of the plots worksheet
'===============================================================================
Option Explicit

Private pStepColumn As String
Private pNameColumn As String
Private pDescColumn As String
Private pUnitColumn As String
Private pKeyColumn As String
Private pTargetColumn As String
Private pMinColumn As String
Private pMaxColumn As String
Private pMacroColumn As String
Private pDataStartColumn As String
Private pDataStartRow As Long
Private pNumColumns As Long
Private pNumRows As Long
Private pBatchRowKey As String
Private pStartDateKey As String
Private pPlotsSheetName As String

'===============================================================================
' [SUB] Class_Initialize
'===============================================================================
' Description:
'   Initializes the SpecsCls instance with default column assignments and
'   row counts matching the standard Data worksheet layout.
'
'   Modify these defaults if your Data worksheet has a different structure, or
'   if you want to adapt the system to a different format. The rest of the
'   code relies on these properties to locate data, so they must be set 
'   correctly.
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
    pNumColumns = 5
    pNumRows = 75
    
    pBatchRowKey = "[00:00] Batchnummer"
    pStartDateKey = "[00:00] Startdato"
    
    pPlotsSheetName = "Plots"
End Sub

Public Property Get StepColumn() As String
    StepColumn = pStepColumn
End Property

Public Property Get NameColumn() As String
    NameColumn = pNameColumn
End Property

Public Property Get DescColumn() As String
    DescColumn = pDescColumn
End Property

Public Property Get UnitColumn() As String
    UnitColumn = pUnitColumn
End Property

Public Property Get KeyColumn() As String
    KeyColumn = pKeyColumn
End Property

Public Property Get TargetColumn() As String
    TargetColumn = pTargetColumn
End Property

Public Property Get MinColumn() As String
    MinColumn = pMinColumn
End Property

Public Property Get MaxColumn() As String
    MaxColumn = pMaxColumn
End Property

Public Property Get MacroColumn() As String
    MacroColumn = pMacroColumn
End Property

Public Property Get DataStartColumn() As String
    DataStartColumn = pDataStartColumn
End Property

Public Property Get DataStartRow() As Long
    DataStartRow = pDataStartRow
End Property

Public Property Get NumColumns() As Long
    NumColumns = pNumColumns
End Property

Public Property Get NumRows() As Long
    NumRows = pNumRows
End Property

Public Property Get BatchRowKey() As String
    BatchRowKey = pBatchRowKey
End Property

Public Property Get StartDateKey() As String
    StartDateKey = pStartDateKey
End Property

Public Property Get PlotsSheetName() As String
    PlotsSheetName = pPlotsSheetName
End Property
