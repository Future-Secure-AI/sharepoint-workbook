[Microsoft Graph SDK](README.md) / Workbook

# Workbook

Workbook models.

## Type Aliases

### Workbook

> **Workbook** = `object`

Defined in: [src/models/Workbook.ts:13](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Workbook.ts#L13)

Represents a workbook with worksheets.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="name"></a> `name` | `string` | Name of the workbook. | [src/models/Workbook.ts:14](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Workbook.ts#L14) |
| <a id="worksheets"></a> `worksheets` | [`Iterable`](https://www.typescriptlang.org/docs/handbook/iterators-and-generators.html#iterable-interface)\<[`Worksheet`](Worksheet.md#worksheet)\> \| `AsyncIterable`\<[`Worksheet`](Worksheet.md#worksheet)\> | Worksheets in the workbook. | [src/models/Workbook.ts:15](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Workbook.ts#L15) |

***

### WriteWorkbook

> **WriteWorkbook** = `object`

Defined in: [src/models/Workbook.ts:23](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Workbook.ts#L23)

Represents a workbook to be written, with writeable worksheets.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="name-1"></a> `name` | `string` | Name of the workbook. | [src/models/Workbook.ts:24](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Workbook.ts#L24) |
| <a id="worksheets-1"></a> `worksheets` | [`Iterable`](https://www.typescriptlang.org/docs/handbook/iterators-and-generators.html#iterable-interface)\<[`WriteWorksheet`](Worksheet.md#writeworksheet)\> \| `AsyncIterable`\<[`WriteWorksheet`](Worksheet.md#writeworksheet)\> | Worksheets to write. | [src/models/Workbook.ts:25](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Workbook.ts#L25) |
