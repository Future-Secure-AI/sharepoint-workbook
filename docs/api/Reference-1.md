[Microsoft Graph SDK](README.md) / Reference

# Reference

References to one or more cells in a worksheet.

## Type Aliases

### CellRef

> **CellRef** = `` `${ColumnRef}${RowRef}` ``

Defined in: [src/models/Reference.ts:24](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L24)

Represents a cell reference in a worksheet (e.g., "A1").

***

### ColumnRef

> **ColumnRef** = `` `${Letter}` ``

Defined in: [src/models/Reference.ts:12](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L12)

Represents a column in a worksheet.

#### Remarks

Only the first columns are covered by TypeScript type checking due to TypeScript complexity limitations, however all columns are supported at runtime.

***

### ExplicitRef

> **ExplicitRef** = `` `${CellRef}:${CellRef}` `` \| \[[`CellRef`](#cellref), [`CellRef`](#cellref)\]

Defined in: [src/models/Reference.ts:43](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L43)

Represents an explicit cell range reference (e.g., "A1:B2" or a tuple).

***

### Letter

> **Letter** = `"A"` \| `"B"` \| `"C"` \| `"D"` \| `"E"` \| `"F"` \| `"G"` \| `"H"` \| `"I"` \| `"J"` \| `"K"` \| `"L"` \| `"M"` \| `"N"` \| `"O"` \| `"P"` \| `"Q"` \| `"R"` \| `"S"` \| `"T"` \| `"U"` \| `"V"` \| `"W"` \| `"X"` \| `"Y"` \| `"Z"`

Defined in: [src/models/Reference.ts:49](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L49)

Represents a single uppercase column letter (A-Z).

***

### Ref

> **Ref** = `ColumnOrRowOrCell` \| \`$\{ColumnOrRowOrCell \| ""\}:$\{ColumnOrRowOrCell \| ""\}\` \| \[`ColumnOrRowOrCell` \| `null`, `ColumnOrRowOrCell` \| `null`\]

Defined in: [src/models/Reference.ts:37](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L37)

Represents a worksheet reference, which can be a single column, row, cell, a range string, or a tuple.

***

### RowRef

> **RowRef** = `` `${number}` ``

Defined in: [src/models/Reference.ts:18](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Reference.ts#L18)

Represents a row in a worksheet. Can be a string or number.
