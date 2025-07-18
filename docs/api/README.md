# Microsoft Graph SDK

## Errors

| Module | Description |
| ------ | ------ |
| [MissingPathError](MissingPathError.md) | Error thrown when attempting to save a file when it hasn't been "saved as", so no path is known. |

## Models

| Module | Description |
| ------ | ------ |
| [Cell](Cell.md) | Cells and its properties in a worksheet. |
| [Column](Column.md) | Columns. |
| [DeepPartial](DeepPartial.md) | Making all properties of a type optional, recursively. |
| [Handle](Handle.md) | Reference to an opened workbook. |
| [LocalFilePath](LocalFilePath.md) | Local file path. |
| [Reference](Reference-1.md) | References to one or more cells in a worksheet. |
| [Row](Row.md) | Rows. |
| [Worksheet](Worksheet.md) | Worksheets. |

## Services

| Module | Description |
| ------ | ------ |
| [cellReader](cellReader.md) | Utilities for reading values and formatting from worksheet cells. |
| [cellWriter](cellWriter.md) | Utilities for writing values and formatting to worksheet cells. |
| [reference](reference.md) | Utilities for parsing and resolving cell, row, column, and range references in worksheets. |
| [streamParameters](streamParameters.md) | Streaming parameter utilities for workbook operations. |
| [temporaryFile](temporaryFile.md) | Utilities for creating temporary file paths for workbook operations. |

## Tasks

| Module | Description |
| ------ | ------ |
| [clearCells](clearCells.md) | Clear all values and formatting in a specified range of cells in a worksheet. |
| [createWorkbook](createWorkbook.md) | Create a new workbook, optionally with specified worksheets. |
| [deleteCells](deleteCells.md) | Deletes a given set of columns or rows from a worksheet. |
| [filterWorkbookColumns](filterWorkbookColumns.md) | Filter out unwanted columns from a workbook. |
| [filterWorkbookRows](filterWorkbookRows.md) | Filter out unwanted rows from a workbook. |
| [findWorksheet](findWorksheet.md) | Find a worksheet in a workbook by name. |
| [getWorksheet](getWorksheet.md) | Get a worksheet from a workbook by its exact name. |
| [insertCells](insertCells.md) | Insert a rectangular block of cells into a worksheet, shifting existing cells down or right. |
| [openWorkbookByPath](openWorkbookByPath.md) | Reading a workbook from SharePoint by path. |
| [readCells](readCells.md) | Read a rectangular block of cells from a worksheet. |
| [readCellValues](readCellValues.md) | Read a rectangular block of cell values from a worksheet (no styles included). |
| [saveWorkbook](saveWorkbook.md) | Write opened workbook back to Microsoft SharePoint. |
| [saveWorkbookAs](saveWorkbookAs.md) | Write workbook to Microsoft Sharepoint to a specific path. |
| [updateCells](updateCells.md) | Update a rectangular block of cells in a worksheet, starting at the given origin. |
| [updateEachCell](updateEachCell.md) | Update every cell in a rectangular range to the same value or partial cell object. |
