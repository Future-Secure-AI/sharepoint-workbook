[Microsoft Graph SDK](README.md) / readCells

# readCells

Read a rectangular block of cells from a worksheet.

## Functions

### readCells()

> **readCells**(`worksheet`, `range`): [`Cell`](Cell.md#cell)[][]

Defined in: [src/tasks/readCells.ts:19](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/readCells.ts#L19)

Reads a rectangular block of cells from the worksheet.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheet` | `Worksheet` | The worksheet to read from. |
| `range` | [`Ref`](Reference-1.md#ref) | The range reference (e.g., "A1:B2") specifying the block to read. |

#### Returns

[`Cell`](Cell.md#cell)[][]

A 2D array of Cell objects representing the values in the specified range.
