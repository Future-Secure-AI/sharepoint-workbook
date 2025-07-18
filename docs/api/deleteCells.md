[Microsoft Graph SDK](README.md) / deleteCells

# deleteCells

Deletes a given set of columns or rows from a worksheet.

## Type Aliases

### ColumnOrRowRangeRef

> **ColumnOrRowRangeRef** = \`$\{ColumnRef \| ""\}:$\{ColumnRef \| ""\}\` \| \`$\{RowRef \| ""\}:$\{RowRef \| ""\}\` \| \[[`ColumnRef`](Reference-1.md#columnref) \| `null`, [`ColumnRef`](Reference-1.md#columnref) \| `null`\] \| \[[`RowRef`](Reference-1.md#rowref) \| `null`, [`RowRef`](Reference-1.md#rowref) \| `null`\]

Defined in: [src/tasks/deleteCells.ts:12](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/deleteCells.ts#L12)

## Functions

### deleteCells()

> **deleteCells**(`worksheet`, `range`): `void`

Defined in: [src/tasks/deleteCells.ts:21](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/deleteCells.ts#L21)

Deletes a given set of columns or rows from a worksheet. Adjacent cells will be shifted up or left.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheet` | `Worksheet` | The worksheet to modify. |
| `range` | [`ColumnOrRowRangeRef`](#columnorrowrangeref) | The range reference (e.g., "A:C" or "1:5") specifying the range to delete. |

#### Returns

`void`

#### Throws

If shift is not "Up" or "Left".
