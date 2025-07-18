[Microsoft Graph SDK](README.md) / updateCells

# updateCells

Update a rectangular block of cells in a worksheet, starting at the given origin.

## Functions

### updateCells()

> **updateCells**(`worksheet`, `origin`, `cells`): `void`

Defined in: [src/tasks/updateCells.ts:21](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/updateCells.ts#L21)

Updates a rectangular block of cells in the worksheet, starting at the given origin.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheet` | `Worksheet` | The worksheet to update. |
| `origin` | `` `Q${number}` `` \| `` `A${number}` `` \| `` `B${number}` `` \| `` `C${number}` `` \| `` `D${number}` `` \| `` `E${number}` `` \| `` `F${number}` `` \| `` `G${number}` `` \| `` `H${number}` `` \| `` `I${number}` `` \| `` `J${number}` `` \| `` `K${number}` `` \| `` `L${number}` `` \| `` `M${number}` `` \| `` `N${number}` `` \| `` `O${number}` `` \| `` `P${number}` `` \| `` `R${number}` `` \| `` `S${number}` `` \| `` `T${number}` `` \| `` `U${number}` `` \| `` `V${number}` `` \| `` `W${number}` `` \| `` `X${number}` `` \| `` `Y${number}` `` \| `` `Z${number}` `` | The top-left cell reference (e.g., "A1") where the update begins. |
| `cells` | (`undefined` \| [`CellValue`](Cell.md#cellvalue-1) \| [`DeepPartial`](DeepPartial.md#deeppartial)\<[`Cell`](Cell.md#cell)\>)[][] | A 2D array of cell values or partial cell objects to write. All rows must have the same length. |

#### Returns

`void`
