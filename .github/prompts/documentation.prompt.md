---
description: "Use when: documenting VBA code, adding function headers, writing procedure documentation, creating docstrings for macros"
---

# VBA Documentation Standards

Document VBA code following NumPy-style documentation conventions adapted for VBA. Documentation must be comprehensive enough for non-technical users to understand.

## Documentation Format

Every Function and Sub must have a header comment block immediately above the procedure definition, using this structure:

```vba
'===============================================================================
' [FUNCTION/MACRO] ProcedureName
'===============================================================================
' Description:
'   A clear, plain-language explanation of what this procedure does and why it
'   exists. Explain its role within the broader project context. Non-technical
'   users should understand the purpose without reading the code.
'
' Parameters:
'   paramName1 : DataType
'       Description of the parameter and its expected values/format.
'   paramName2 : DataType, Optional
'       Description. Default: [default value]
'
' Returns:
'   DataType
'       Description of what is returned and its meaning.
'       (Omit this section for Sub procedures / Macros)
'
' Example:
'   result = ProcedureName(arg1, arg2)
'
' Notes:
'   - Any important caveats, limitations, or side effects
'   - Dependencies on other procedures or external data
'===============================================================================
```

## Requirements

1. **Procedure Type**: Always specify `[FUNCTION]` (a function or method function with arguments and a return value), `[SUB]` (a sub with input parameters but no return value, or a method sub with no return value), `[MACRO]` (a non-method sub with no input parameters and no return value) in the header, or `[PROPERTY]` (a property procedure with a return value) in the header. This allows users to quickly understand the nature of the procedure at a glance.
2. **Description**: Analyse the entire project to understand how this procedure fits into the overall workflow. Explain its purpose in context.
3. **Parameters**: List ALL parameters with their data types and clear descriptions
4. **Returns**: For Functions, describe the return value and its meaning
5. **Plain Language**: Write for non-technical users — avoid jargon, explain Excel/VBA concepts when relevant
6. **Example**: Include a usage example when the procedure's usage isn't immediately obvious

## Context

Before documenting, examine the project structure in `src/magle_vba/` to understand:
- How the procedure relates to other modules
- What data flows through it
- Its role in the user's workflow

## Task

{{task}}
