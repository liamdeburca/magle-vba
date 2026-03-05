# magle-vba

## Introduction

The goal of this tutorial is to give the user a beginner-friendly overview of the template -- how to use it to create new Excel spreadsheets and how to use and manage the in-built macros and functions. 

In general, a good rule-of-thumb: do NOT change the current format (name of the "Data" / "Plots" sheets, number of header rows, layout of columns etc.). The current format is well-tested. 

## The "Data" Sheet

The worksheet named "Data" is supposed to be the main worksheet of the user -- where they input their data and perform most of their preliminary data analysis. All backend code relies on the expected layout of this sheet, so the format should not be changed without also updating the backend code. 

### Step, Name, Description, and Key 

The first three columns describe the full title of a tracked subprocess:
1. The subprocess step with the format: "XX:YY"
2. The subprocess name.
3. (Optional) The subprocess description.

The combination of all three descriptors creates the "key" value for the row. For defined behaviour, this key should be unique -- if two rows have the same steps and names, use the description to discriminate between the two rows. The "key" value is written to the column labelled "Key".

When the code tries to parse the "Data" worksheet, it looks for rows with non-empty step and name values. Once a row has been previously found, parsing will terminate once both values are empty. It is therefore useful to collect all subprocesses in consecutive rows, or separate groups of subprocesses by a row with filler text in the step/name cell(s). 

> Note: in code, each row can either be referenced using the row index or the "key" value written to the "Key" column. 

### The Unit

The unit is mostly for show, but still specifies how to parse the data in the given row. There are four possibilities:

1. hh:mm - Convert to time-values.
2. åååå-mm-dd OR "åååå-mm" - Convert to date-value.
3. txt - No conversion, best suited for values with inconsistent format or no analytical use such as week number.  
4. MISSING OR kg OR L etc. - In all other cases values are assumed to be numeric. 

> Note: In the future, the unit-value may be made obligatory and used for data validation and formatting of the entire row. 

### Target, Min, and Max

The values in these columns are optional. They are used in data formatting and visualisation pipelines. These values should come from the subprocess' journal. I suggest that in the case that no lower and upper bounds are given in the journal, the user uses their intuition wherever possible -- percentages shouldn't be negative or >100% etc. 

### Macro

This column is initially empty. Once the "Start" macro is run, this cell is populated with a drop-down menu of possible macros to run -- these include sorting and visualisation utilities. Once a macro is selected and it runs successfully, the cell should be automatically cleared. If this doesn't happen, the macro did not run completely.

### Raw Data

All subsequent rows should contain the raw data associated with each subprocess. The backend code is not intelligent enough to interpret non-standard data, so input values should make sense:

1. If a subprocess' parameter is inherently numeric, don't input text as this will likely be ignored, or cause an error the user then has to identify. 
2. If a subprocess' parameter is constant for all processes, the user should still input these constant values. Missing values are interpreted as missing. 

> Note: In the future, one may implement a "default" value for each row. Missing values are automatically replaced with this "default" value. 

## The "Plots" Sheet

The "Plots" worksheet has no functional value other than the default output location of figures created by the visualisation macros. The user can change the name of this sheet, but if they do, they must also update the `pPlotsSheetName` property in the `SpecsCls` class module.

## The "README" Sheet

The "README" worksheet contains a detailed list of all macros with the possibility to run them directly from the sheet. This is meant to be a user-friendly interface for non-technical users to interact with the macros without having to open the VBA editor.

# Backend Code - How It All Works

## The "SpecsCls" Class Module

The `SpecsCls` class module contains all the specifications for how to parse the "Data" worksheet. This includes which columns to look for specific values (e.g. steps, names, units, macros etc.) and how many header rows to skip when parsing. The `Class_Initialize` method sets default values for these specifications, which can be modified if the user has a different layout in their "Data" worksheet.

It has the following attributes:
- `pStepColumn` (String): The column letter where the subprocess steps are located. Default: "A".
- `pNameColumn` (String): The column letter where the subprocess names are located. Default: "B".
- `pDescColumn` (String): The column letter where the subprocess descriptions are located. Default: "C".
- `pUnitColumn` (String): The column letter where the units are located. Default: "D".
- `pKeyColumn` (String): The column letter where the "key" values are located. Default: "E". 
- `pTargetColumn` (String): The column letter where the target values are located. Default: "F".
- `pMinColumn` (String): The column letter where the minimum values are located. Default: "G".
- `pMaxColumn` (String): The column letter where the maximum values are located. Default: "H".
- `pMacroColumn` (String): The column letter where the macro drop-down menu is located. Default: "I".
- `pDataStartColumn` (String): The column letter where the raw data starts. Default: "J".
- `pDataStartRow` (Long): The row number where the raw data starts. Default: 3.
- `pNumColumns` (Long): The total number of columns to parse in the "Data" worksheet. Adjust this to maximise computational efficiency. Default: 1000.
- `pNumRows` (Long): The total number of rows to parse in the "Data" worksheet if the parsing does not terminate by itself (see Section "Step, Name, Description, and Key"). Adjust this to maximise computational efficiency. Default: 1000.
- `pBatchRowKey` (String): The key to the row containing the batch identifier (number). This is used as the primary (x) axis in some standard plotting utilities. Default: "[00:00] Batchnummer". 
- `pStartDateKey` (String): The key to the row containing the start date for each batch. This is used as the secondary x-axis in some standard plotting utilities. Default: "[00:00] Startdato".
- `pPlotsName` (String): The name of the worksheet where some standard plotting utilities should create their figures. Default: "Plots".

## The "DataRowCls" Class Module

The `DataRowCls` class module contains raw and parsed values for a specific row. This includes the following:

- `pStep` (String): The subprocess step identifier in "XX:YY" format. Default: Empty string.
- `pName` (String): The subprocess name. Default: Empty string.
- `pDesc` (String): The subprocess description. Default: Empty string.
- `pUnit` (String): The unit of measurement or data type specifier. Default: Empty string.
- `pTarget` (Variant): The target value for the subprocess parameter. Default: #N/A (xlErrNA).
- `pMin` (Variant): The minimum acceptable value for the subprocess parameter. Default: #N/A (xlErrNA).
- `pMax` (Variant): The maximum acceptable value for the subprocess parameter. Default: #N/A (xlErrNA).
- `pRowIdx` (Long): The row index of this row in the "Data" worksheet. Default: 0.
- `pRawData` (Variant array): An array containing all raw cell values for the subprocess across batches. Default: Empty array.
- `pDblData` (Variant array): An array containing values converted to Double where possible. Default: Empty array.
- `pTxtData` (String array): An array containing values as text strings. Default: Empty array.
- `pData` (Variant array): An array containing parsed/converted data values according to the unit specification. Default: Empty array.

The `key` property is computed from the combination of step, name, and description (not stored as a private attribute):

```vba
Public Property Get key() As String
    key = "[" & Me.pStep & "] " & Me.pName
    If Me.pDesc <> "" Then
        key = key & " - " & Me.pDesc
    End If
End Property
```

## The "ParsedDataCls" Class Module

The `ParsedDataCls` class module contains a collection of `DataRowCls` instances, each representing a row in the "Data" worksheet. It also contains methods for sorting the rows based on the data values of a specific row (e.g. sorting all rows based on the batch numbers in the row with key "[00:00] Batchnummer").

It has the following attributes:

- `pRows` (Collection): A collection of DataRowCls instances, each representing a row from the Data worksheet. Default: Empty collection (initialised in Class_Initialize).
- `pNumColumnsMax` (Long): The maximum number of non-empty data columns across all rows, determined during the Crop operation. Default: 0.
- `pNumRowsMax` (Long): The maximum number of rows loaded from the Data worksheet. Default: 0.

It has the following properties:

- `Rows` (Collection): Returns the pRows collection containing all DataRowCls instances.
- `count` (Long): Returns the number of DataRowCls instances in the collection (pRows.count).
- `Keys` (String array): Returns an array of all row keys for lookup purposes.