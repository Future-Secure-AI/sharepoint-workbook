[Microsoft Graph SDK](../README.md) / services/reference

# services/reference

## Functions

### columnComponentToNumber()

> **columnComponentToNumber**(`column`): `number`

Defined in: [src/services/reference.ts:75](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L75)

Converts a column reference (e.g., "A", "Z") to its number (1-based).

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `column` | [`Letter`](../models/Reference.md#letter) |

#### Returns

`number`

***

### parseCellReference()

> **parseCellReference**(`cell`): \[`number`, `number`\]

Defined in: [src/services/reference.ts:17](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L17)

Parses a cell reference (e.g., "A1") into [col, row] numbers.

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `cell` | `` `Q${number}` `` \| `` `A${number}` `` \| `` `B${number}` `` \| `` `C${number}` `` \| `` `D${number}` `` \| `` `E${number}` `` \| `` `F${number}` `` \| `` `G${number}` `` \| `` `H${number}` `` \| `` `I${number}` `` \| `` `J${number}` `` \| `` `K${number}` `` \| `` `L${number}` `` \| `` `M${number}` `` \| `` `N${number}` `` \| `` `O${number}` `` \| `` `P${number}` `` \| `` `R${number}` `` \| `` `S${number}` `` \| `` `T${number}` `` \| `` `U${number}` `` \| `` `V${number}` `` \| `` `W${number}` `` \| `` `X${number}` `` \| `` `Y${number}` `` \| `` `Z${number}` `` |

#### Returns

\[`number`, `number`\]

***

### parseRangeReference()

> **parseRangeReference**(`range`): \[`null` \| `number`, `null` \| `number`, `null` \| `number`, `null` \| `number`\]

Defined in: [src/services/reference.ts:28](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L28)

Converts a RangeRef to an array: [startCol, startRow, endCol, endRow].

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `range` | [`RangeRef`](../models/Reference.md#rangeref) | RangeRef (array or string) |

#### Returns

\[`null` \| `number`, `null` \| `number`, `null` \| `number`, `null` \| `number`\]

[startCol, startRow, endCol, endRow]

***

### parseRangeReferenceExact()

> **parseRangeReferenceExact**(`range`, `worksheet`): \[`number`, `number`, `number`, `number`\]

Defined in: [src/services/reference.ts:61](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L61)

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `range` | [`RangeRef`](../models/Reference.md#rangeref) |
| `worksheet` | `Worksheet` |

#### Returns

\[`number`, `number`, `number`, `number`\]

***

### rowComponentToNumber()

> **rowComponentToNumber**(`row`): `number`

Defined in: [src/services/reference.ts:86](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/reference.ts#L86)

Converts a row reference (string or number) to a number.

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `row` | [`RowRef`](../models/Reference.md#rowref) |

#### Returns

`number`
