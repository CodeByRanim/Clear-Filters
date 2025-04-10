# Clear Filters in All Sheets

This macro loops through all visible sheets in the active workbook and clears any active filters. It's helpful when working with shared files where filters might be left applied across different sheets, causing confusion or missing data.

## Features

- Clears all filters in visible sheets
- Leaves hidden sheets untouched
- Works instantly with a single macro

## How to Use

1. Open the VBA Editor (ALT + F11)
2. Insert a new module
3. Paste the code from `ClearAllFilters.bas`
4. Run the macro `ClearAllFilters`

## Example

Before:
- Sheet1: Filter applied on column A  
- Sheet2: Filter applied on column B  
- Sheet3: Hidden  

After running the macro:
- Filters removed from Sheet1 and Sheet2  
- Sheet3 unchanged
