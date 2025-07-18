[Microsoft Graph SDK](README.md) / createWorkbook

# createWorkbook

Create a new workbook, optionally with specified worksheets.

## Functions

### createWorkbook()

> **createWorkbook**(`worksheets?`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Workbook`](Handle.md#workbook)\>

Defined in: [src/tasks/createWorkbook.ts:31](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/createWorkbook.ts#L31)

Create a new workbook, optionally with specified worksheets.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheets?` | [`Record`](https://www.typescriptlang.org/docs/handbook/utility-types.html#recordkeys-type)\<`string`, ([`CellValue`](Cell.md#cellvalue-1) \| [`DeepPartial`](DeepPartial.md#deeppartial)\<[`Cell`](Cell.md#cell)\>)[][]\> | An object whose keys are worksheet names and values are iterables or async iterables of row values. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Workbook`](Handle.md#workbook)\>

Handle referencing the newly created workbook.

#### Example

```ts
const handle = await createWorkbook({
  Sheet1: [
    [1, 2, 3],
    [4, 5, 6],
  ],
  Sheet2: [
    ["A", "B", "C"],
    ["D", "E", "F"],
  ],
});
```
