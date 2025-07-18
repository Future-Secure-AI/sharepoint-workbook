[Microsoft Graph SDK](README.md) / reference

# reference

Utilities for parsing and resolving cell, row, column, and range references in worksheets.

## Functions

### parseCellRef()

> **parseCellRef**(`cell`): \[[`ColumnIndex`](Column.md#columnindex), [`RowIndex`](Row.md#rowindex)\]

Defined in: [src/services/reference.ts:25](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L25)

Parses a cell reference (e.g., "A1") into a tuple of 0-based column and row indices.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `cell` | `` `Q${number}` `` \| `` `A${number}` `` \| `` `B${number}` `` \| `` `C${number}` `` \| `` `D${number}` `` \| `` `E${number}` `` \| `` `F${number}` `` \| `` `G${number}` `` \| `` `H${number}` `` \| `` `I${number}` `` \| `` `J${number}` `` \| `` `K${number}` `` \| `` `L${number}` `` \| `` `M${number}` `` \| `` `N${number}` `` \| `` `O${number}` `` \| `` `P${number}` `` \| `` `R${number}` `` \| `` `S${number}` `` \| `` `T${number}` `` \| `` `U${number}` `` \| `` `V${number}` `` \| `` `W${number}` `` \| `` `X${number}` `` \| `` `Y${number}` `` \| `` `Z${number}` `` | Reference string such as "B2". |

#### Returns

\[[`ColumnIndex`](Column.md#columnindex), [`RowIndex`](Row.md#rowindex)\]

0-based column and row indices.

#### Throws

If the cell reference format is invalid.

***

### parseRef()

> **parseRef**(`range`): \[`null` \| [`ColumnIndex`](Column.md#columnindex), `null` \| [`RowIndex`](Row.md#rowindex), `null` \| [`ColumnIndex`](Column.md#columnindex), `null` \| [`RowIndex`](Row.md#rowindex)\]

Defined in: [src/services/reference.ts:38](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L38)

Converts a range reference (string or array) to a tuple of 0-based indices: [startCol, startRow, endCol, endRow].
Accepts cell, row, column, or range references in string or array form.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `range` | [`Ref`](Reference-1.md#ref) | Reference such as "A1:C3", ["A", "C"], [1, 5], etc. |

#### Returns

\[`null` \| [`ColumnIndex`](Column.md#columnindex), `null` \| [`RowIndex`](Row.md#rowindex), `null` \| [`ColumnIndex`](Column.md#columnindex), `null` \| [`RowIndex`](Row.md#rowindex)\]

Resolved indices (null if not specified).

#### Throws

If the reference is invalid or the range ends before it starts.

***

### parseRefResolved()

> **parseRefResolved**(`range`, `worksheet`): \[[`ColumnIndex`](Column.md#columnindex), [`RowIndex`](Row.md#rowindex), [`ColumnIndex`](Column.md#columnindex), [`RowIndex`](Row.md#rowindex)\]

Defined in: [src/services/reference.ts:130](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L130)

Resolves a range reference to concrete 0-based indices, filling in worksheet bounds for omitted values.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `range` | [`Ref`](Reference-1.md#ref) | Reference to resolve. |
| `worksheet` | `Worksheet` | Worksheet to use for bounds. |

#### Returns

\[[`ColumnIndex`](Column.md#columnindex), [`RowIndex`](Row.md#rowindex), [`ColumnIndex`](Column.md#columnindex), [`RowIndex`](Row.md#rowindex)\]

Fully resolved indices: [startCol, startRow, endCol, endRow].

***

### resolveColumnIndex()

> **resolveColumnIndex**(`column`): [`ColumnIndex`](Column.md#columnindex)

Defined in: [src/services/reference.ts:146](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L146)

Converts a column reference (e.g., "A", "Z", "AA") to its 0-based column index.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `column` | [`Letter`](Reference-1.md#letter) | Column reference string. |

#### Returns

[`ColumnIndex`](Column.md#columnindex)

0-based column index.

***

### resolveRowIndex()

> **resolveRowIndex**(`row`): [`RowIndex`](Row.md#rowindex)

Defined in: [src/services/reference.ts:160](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L160)

Converts a row reference (string or number) to a 0-based row index.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `row` | [`RowRef`](Reference-1.md#rowref) | Row reference (number or string). |

#### Returns

[`RowIndex`](Row.md#rowindex)

0-based row index.

#### Throws

If the row reference is invalid.
