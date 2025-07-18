[Microsoft Graph SDK](README.md) / updateEachCell

# updateEachCell

Applies an update to every cell in the specified range of a worksheet.

## Functions

### updateEachCell()

> **updateEachCell**(`worksheet`, `range`, `write`): `void`

Defined in: [src/tasks/updateEachCell.ts:25](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/updateEachCell.ts#L25)

Applies an update to every cell in the specified range of a worksheet.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheet` | `Worksheet` | The worksheet to update. |
| `range` | [`Ref`](Reference-1.md#ref) | The range reference (e.g., "A1:B2") specifying the block to update. |
| `write` | [`CellValue`](Cell.md#cellvalue-1) \| [`DeepPartial`](DeepPartial.md#deeppartial)\<[`Cell`](Cell.md#cell)\> | The value or partial cell object to write to each cell in the range. |

#### Returns

`void`

#### Example

```ts
// Updates every cell in the range A1:B2 to have a value of 42
updateEachCell(worksheet, "A1:B2", 42);

// Updates every cell in the first row to be bold
updateEachCell(worksheet, "1", { fontBold: true });
```
