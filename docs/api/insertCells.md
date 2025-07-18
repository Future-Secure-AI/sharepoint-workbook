[Microsoft Graph SDK](README.md) / insertCells

# insertCells

Insert a rectangular block of cells into a worksheet, shifting existing cells down or right.

## Functions

### insertCells()

> **insertCells**(`worksheet`, `origin`, `shift`, `cells`): `void`

Defined in: [src/tasks/insertCells.ts:23](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/insertCells.ts#L23)

Inserts a rectangular block of cells into the worksheet, shifting existing cells either down or right.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheet` | `Worksheet` | The worksheet to modify. |
| `origin` | `` `Q${number}` `` \| `` `A${number}` `` \| `` `B${number}` `` \| `` `C${number}` `` \| `` `D${number}` `` \| `` `E${number}` `` \| `` `F${number}` `` \| `` `G${number}` `` \| `` `H${number}` `` \| `` `I${number}` `` \| `` `J${number}` `` \| `` `K${number}` `` \| `` `L${number}` `` \| `` `M${number}` `` \| `` `N${number}` `` \| `` `O${number}` `` \| `` `P${number}` `` \| `` `R${number}` `` \| `` `S${number}` `` \| `` `T${number}` `` \| `` `U${number}` `` \| `` `V${number}` `` \| `` `W${number}` `` \| `` `X${number}` `` \| `` `Y${number}` `` \| `` `Z${number}` `` | The top-left cell reference (e.g., "A1") where the insertion begins. |
| `shift` | `"right"` \| `"down"` | The direction to shift existing cells: "down" or "right". |
| `cells` | ([`CellValue`](Cell.md#cellvalue-1) \| [`DeepPartial`](DeepPartial.md#deeppartial)\<[`Cell`](Cell.md#cell)\>)[][] | A 2D rectangular array of cell values or partial cell objects to insert. All rows must have the same length. |

#### Returns

`void`

#### Throws

If rows in `cells` have different lengths, or if `shift` is not "Down" or "Right".
