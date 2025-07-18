[Microsoft Graph SDK](README.md) / cellReader

# cellReader

Utilities for reading values and formatting from worksheet cells.

## Functions

### readCell()

> **readCell**(`worksheet`, `r`, `c`): [`Cell`](Cell.md#cell)

Defined in: [src/services/cellReader.ts:25](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/cellReader.ts#L25)

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `worksheet` | `Worksheet` |
| `r` | [`RowIndex`](Row.md#rowindex) |
| `c` | [`ColumnIndex`](Column.md#columnindex) |

#### Returns

[`Cell`](Cell.md#cell)

***

### readCellValue()

> **readCellValue**(`worksheet`, `r`, `c`): [`CellValue`](Cell.md#cellvalue-1)

Defined in: [src/services/cellReader.ts:19](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/cellReader.ts#L19)

Reads the value of a cell from a worksheet.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheet` | `Worksheet` | Worksheet instance. |
| `r` | [`RowIndex`](Row.md#rowindex) | Row index (0-based). |
| `c` | [`ColumnIndex`](Column.md#columnindex) | Column index (0-based). |

#### Returns

[`CellValue`](Cell.md#cellvalue-1)

Value of the cell (string, number, boolean, Date, or empty string).
