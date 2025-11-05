# Excels — VBA Toolkit

Reusable VBA modules and classes for Excel automation. The repo stores code as text files (`.bas`, `.cls`, `.frm`) so you can diff and review changes.

## Structure
src/
Modules/ ' .bas files
Classes/ ' .cls files
Forms/ ' .frm and .frx (binary) files
examples/ ' optional demo workbooks kept for reference

## Importing modules into Excel
1. Open Excel. Enable the **Developer** tab if needed.
2. Developer → **Visual Basic**.
3. In the VBA editor: **File → Import File…** and pick the `.bas` or `.cls` from `src`.
4. Save your workbook as `.xlsm` or `.xlam`.

## Exporting modules from Excel
In the VBA editor: right-click a module or class → **Export File…**. Save into the matching `src/` subfolder.

## Conventions
- One public routine per file when possible.
- `Option Explicit` at the top of every code file.
- PascalCase for public procedures. camelCase for locals.
- No hard-coded worksheet names in shared modules. Pass them as parameters.

## Example module
Create `src/Modules/DebugPrint.bas`:
```vb
Attribute VB_Name = "DebugPrint"
Option Explicit

' Prints to the Immediate Window with a timestamp.
Public Sub LogLine(ByVal message As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"); "  "; message
End Sub
