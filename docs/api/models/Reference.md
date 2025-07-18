[Microsoft Graph SDK](../README.md) / models/Reference

# models/Reference

## Type Aliases

### CellRef

> **CellRef** = `` `${ColumnRef}${RowRef}` ``

Defined in: [src/models/Reference.ts:9](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L9)

***

### ColumnRef

> **ColumnRef** = `` `${Letter}` ``

Defined in: [src/models/Reference.ts:7](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L7)

Represents a column in a worksheet.

#### Remarks

Only the first columns are covered by TypeScript type checking due to TypeScript complexity limitations, however all columns are supported at runtime.

***

### ExplicitRef

> **ExplicitRef** = `` `${CellRef}:${CellRef}` `` \| \[[`CellRef`](#cellref), [`CellRef`](#cellref)\]

Defined in: [src/models/Reference.ts:13](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L13)

***

### Letter

> **Letter** = `"A"` \| `"B"` \| `"C"` \| `"D"` \| `"E"` \| `"F"` \| `"G"` \| `"H"` \| `"I"` \| `"J"` \| `"K"` \| `"L"` \| `"M"` \| `"N"` \| `"O"` \| `"P"` \| `"Q"` \| `"R"` \| `"S"` \| `"T"` \| `"U"` \| `"V"` \| `"W"` \| `"X"` \| `"Y"` \| `"Z"`

Defined in: [src/models/Reference.ts:1](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L1)

***

### Ref

> **Ref** = `ColumnRowOrCell` \| \`$\{ColumnRowOrCell \| ""\}:$\{ColumnRowOrCell \| ""\}\` \| \[`ColumnRowOrCell` \| `null`, `ColumnRowOrCell` \| `null`\]

Defined in: [src/models/Reference.ts:12](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L12)

***

### RowRef

> **RowRef** = `` `${number}` `` \| `number`

Defined in: [src/models/Reference.ts:8](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L8)
