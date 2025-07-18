[Microsoft Graph SDK](README.md) / readCellValues

# readCellValues

Read a rectangular block of cell values from a worksheet (no styles included).

## Functions

### readCellValues()

> **readCellValues**(`worksheet`, `range`): [`CellValue`](models/Cell.md#cellvalue-1)[][]

Defined in: [src/tasks/readCellValues.ts:19](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/readCellValues.ts#L19)

Reads a rectangular block of cell values from the worksheet. No styles are included.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheet` | `Worksheet` | The worksheet to read from. |
| `range` | [`Ref`](models/Reference.md#ref) | The range reference (e.g., "A1:B2") specifying the block to read. |

#### Returns

[`CellValue`](models/Cell.md#cellvalue-1)[][]

A 2D array of CellValue objects representing the values in the specified range.
