# magle-vba

A VBA package for Microsoft Excel that provides reusable classes and modules for parsing spreadsheet data for analysis and visualization.

## Overview

`magle-vba` is intended to be imported into an Excel workbook (or the VBA editor) and used as a small toolkit for:

- Reading and parsing worksheet/tabular data
- Normalizing and transforming values into analysis-friendly structures
- Supporting downstream analysis and visualization workflows

## Features

- VBA modules and class modules designed for use in Excel
- Helpers for working with ranges, tables, and worksheet data (as applicable)
- Lightweight, dependency-free (Excel/VBA only)

> Note: Feature details may evolve as modules/classes are added.

## Requirements

- Microsoft Excel (Windows) with VBA enabled
- Macro-enabled workbook support (`.xlsm`) recommended

## Installation

1. Download or clone this repository.
2. In Excel, open the VBA Editor (`ALT + F11`).
3. Import the modules/classes:
   - **File → Import File…**
   - Select the `.bas`, `.cls`, and/or `.frm` files from this repo
4. Save your workbook as a macro-enabled workbook (`.xlsm`).

## Usage

1. Open the Excel workbook where you imported the modules/classes.
2. Call the relevant procedures/functions from your own VBA code.

```vb
' Example (placeholder)
' Sub Example()
'     ' Use magle-vba helpers/classes here
' End Sub
```

## Project structure

- `README.md` — Project documentation
- VBA source files (`.bas`, `.cls`) — Modules and class modules (see repository tree)

## Contributing

Contributions are welcome.

- Open an issue to discuss bugs, enhancements, or ideas.
- Submit a pull request with clear description of changes.

## License

Add a license file (e.g., MIT) and reference it here.

## Changelog

If you maintain releases, consider adding a `CHANGELOG.md` and recording notable changes here.