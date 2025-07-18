[Microsoft Graph SDK](README.md) / updateEachCell

# updateEachCell

Update every cell in a rectangular range to the same value or partial cell object.

## Functions

### updateEachCell()

> **updateEachCell**(`worksheet`, `range`, `write`): `void`

Defined in: [src/tasks/updateEachCell.ts:19](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/updateEachCell.ts#L19)

Updates every cell in the specified rectangular range to the given value or partial cell object.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheet` | `Worksheet` | The worksheet to update. |
| `range` | [`Ref`](models/Reference.md#ref) | The range reference (e.g., "A1:B2") specifying the block to update. |
| `write` | [`CellValue`](models/Cell.md#cellvalue-1) \| [`DeepPartial`](models/DeepPartial.md#deeppartial)\<[`Cell`](models/Cell.md#cell)\> | The value or partial cell object to write to each cell in the range. |

#### Returns

`void`
