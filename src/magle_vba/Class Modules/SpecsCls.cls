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
