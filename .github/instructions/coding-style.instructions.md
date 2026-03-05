---
applyTo: "**/*.bas,**/*.cls"
description: "VBA coding style conventions for this project"
---

# VBA Coding Style

Follow these coding conventions for all VBA code in this project.

## Line Length

- Maximum line length: **80 characters**
- Use line continuation (`_`) to break long lines

## Function/Sub Definitions

Each parameter must appear on its own line:

```vba
Private Function PrepareData( _
    xDataRow As DataRowCls, _
    yDataRow As DataRowCls, _
    Optional mask As Variant, _
    Optional yConversion As Double = 1# _
) As Collection
```

## Comments

- **Minimize in-code comments** — code should be self-documenting
- Use procedure header documentation (see `/documentation` prompt) instead of inline comments
- Only add inline comments for genuinely non-obvious logic

## Variable Declarations

- **Declare variables close to first use**, NOT at the top of the procedure
- This improves readability and keeps related code together

Correct:
```vba
Sub ProcessData()
    Dim inputRange As Range
    Set inputRange = Selection
    
    ' ... some code using inputRange ...
    
    Dim outputArray() As Double
    ReDim outputArray(1 To 10)
    ' ... code using outputArray ...
End Sub
```

Incorrect:
```vba
Sub ProcessData()
    Dim inputRange As Range
    Dim outputArray() As Double
    
    Set inputRange = Selection
    ' ... many lines later ...
    ReDim outputArray(1 To 10)
End Sub
```
