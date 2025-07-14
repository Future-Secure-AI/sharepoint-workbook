[Microsoft Graph SDK](README.md) / importWorkbook

# importWorkbook

Imports worksheet content as a new open workbook.

## Functions

### importWorkbook()

> **importWorkbook**(`worksheets`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Handle`](Handle.md#handle)\>

Defined in: [src/tasks/importWorkbook.ts:19](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/importWorkbook.ts#L19)

Imports worksheet content as a new open workbook.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `worksheets` | [`Iterable`](https://www.typescriptlang.org/docs/handbook/iterators-and-generators.html#iterable-interface)\<[`WriteWorksheet`](Worksheet.md#writeworksheet), `any`, `any`\> \| `AsyncIterable`\<[`WriteWorksheet`](Worksheet.md#writeworksheet), `any`, `any`\> | Worksheet data to import. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Handle`](Handle.md#handle)\>

Handle referencing the newly created workbook.
