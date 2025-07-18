# Microsoft Graph SDK

## Errors

| Module | Description |
| ------ | ------ |
| [MissingPathError](MissingPathError.md) | Error thrown when attempting to save a file when it hasn't been "saved as", so no path is known. |

## Models

| Module | Description |
| ------ | ------ |
| [Handle](Handle.md) | Reference to an opened workbook. |
| [LocalFilePath](LocalFilePath.md) | Local file path. |
| [OptimizeOptions](OptimizeOptions.md) | Options for optimizing a workbook. |
| [Options](Options.md) | Options for operation. |
| [Worksheet](Worksheet.md) | Worksheet models. |

## Other

| Module | Description |
| ------ | ------ |
| [models/Cell](models/Cell.md) | - |
| [models/Column](models/Column.md) | - |
| [models/DeepPartial](models/DeepPartial.md) | - |
| [models/Reference](models/Reference.md) | - |
| [models/Row](models/Row.md) | - |
| [models/Shift](models/Shift.md) | - |
| [services/cellReader](services/cellReader.md) | - |
| [services/cellWriter](services/cellWriter.md) | - |
| [services/rectangularArray](services/rectangularArray.md) | - |
| [services/reference](services/reference.md) | - |
| [services/streamParameters](services/streamParameters.md) | - |
| [services/temporaryFile](services/temporaryFile.md) | - |

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
