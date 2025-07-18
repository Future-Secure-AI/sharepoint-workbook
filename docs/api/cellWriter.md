[Microsoft Graph SDK](README.md) / cellWriter

# cellWriter

Utilities for writing values and formatting to worksheet cells.

## Functions

### writeCell()

> **writeCell**(`worksheet`, `r`, `c`, `cellOrValue`): `void`

Defined in: [src/services/cellWriter.ts:20](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/cellWriter.ts#L20)

Writes a value or cell configuration to a worksheet cell, including formatting, merging, and comments.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheet` | `Worksheet` | Worksheet instance. |
| `r` | [`RowIndex`](Row.md#rowindex) | Row index (0-based). |
| `c` | [`ColumnIndex`](Column.md#columnindex) | Column index (0-based). |
| `cellOrValue` | [`CellValue`](Cell.md#cellvalue-1) \| [`DeepPartial`](DeepPartial.md#deeppartial)\<[`Cell`](Cell.md#cell)\> | Value or partial cell configuration. |

#### Returns

`void`
