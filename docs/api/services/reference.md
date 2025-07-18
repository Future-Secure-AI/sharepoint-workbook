[Microsoft Graph SDK](../README.md) / services/reference

# services/reference

## Functions

### parseCellRef()

> **parseCellRef**(`cell`): \[[`ColumnIndex`](../models/Column.md#columnindex), [`RowIndex`](../models/Row.md#rowindex)\]

Defined in: [src/services/reference.ts:17](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L17)

Parses a cell reference (e.g., "A1") into [col, row] numbers (0-based).

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `cell` | `` `Q${number}` `` \| `` `A${number}` `` \| `` `B${number}` `` \| `` `C${number}` `` \| `` `D${number}` `` \| `` `E${number}` `` \| `` `F${number}` `` \| `` `G${number}` `` \| `` `H${number}` `` \| `` `I${number}` `` \| `` `J${number}` `` \| `` `K${number}` `` \| `` `L${number}` `` \| `` `M${number}` `` \| `` `N${number}` `` \| `` `O${number}` `` \| `` `P${number}` `` \| `` `R${number}` `` \| `` `S${number}` `` \| `` `T${number}` `` \| `` `U${number}` `` \| `` `V${number}` `` \| `` `W${number}` `` \| `` `X${number}` `` \| `` `Y${number}` `` \| `` `Z${number}` `` |

#### Returns

\[[`ColumnIndex`](../models/Column.md#columnindex), [`RowIndex`](../models/Row.md#rowindex)\]

***

### parseRef()

> **parseRef**(`range`): \[`null` \| `number`, `null` \| `number`, `null` \| `number`, `null` \| `number`\]

Defined in: [src/services/reference.ts:28](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L28)

Converts a RangeRef to an array: [startCol, startRow, endCol, endRow] (0-based).

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `range` | [`Ref`](../models/Reference.md#ref) | RangeRef (array or string) |

#### Returns

\[`null` \| `number`, `null` \| `number`, `null` \| `number`, `null` \| `number`\]

[startCol, startRow, endCol, endRow]

***

### parseRefResolved()

> **parseRefResolved**(`range`, `worksheet`): \[`number`, `number`, `number`, `number`\]

Defined in: [src/services/reference.ts:114](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L114)

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `range` | [`Ref`](../models/Reference.md#ref) |
| `worksheet` | `Worksheet` |

#### Returns

\[`number`, `number`, `number`, `number`\]

***

### resolveColumnIndex()

> **resolveColumnIndex**(`column`): [`ColumnIndex`](../models/Column.md#columnindex)

Defined in: [src/services/reference.ts:128](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L128)

Converts a column reference (e.g., "A", "Z") to its number (0-based).

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `column` | [`Letter`](../models/Reference.md#letter) |

#### Returns

[`ColumnIndex`](../models/Column.md#columnindex)

***

### resolveRowIndex()

> **resolveRowIndex**(`row`): [`RowIndex`](../models/Row.md#rowindex)

Defined in: [src/services/reference.ts:139](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L139)

Converts a row reference (string or number) to a number (0-based).

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `row` | [`RowRef`](../models/Reference.md#rowref) |

#### Returns

[`RowIndex`](../models/Row.md#rowindex)
