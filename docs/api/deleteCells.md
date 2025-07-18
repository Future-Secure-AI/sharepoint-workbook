[Microsoft Graph SDK](README.md) / deleteCells

# deleteCells

Delete a rectangular block of cells from a worksheet, shifting remaining cells up or left.

## Functions

### deleteCells()

> **deleteCells**(`worksheet`, `range`, `shift`): `void`

Defined in: [src/tasks/deleteCells.ts:21](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/deleteCells.ts#L21)

Deletes a rectangular block of cells from the worksheet, shifting remaining cells up or left.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheet` | `Worksheet` | The worksheet to modify. |
| `range` | [`RangeRef`](models/Reference.md#rangeref) | The range reference (e.g., "A1:B2") specifying the block to delete. |
| `shift` | [`DeleteShift`](models/Shift.md#deleteshift) | The direction to shift remaining cells: "Up" or "Left". |

#### Returns

`void`

#### Throws

If shift is not "Up" or "Left".
