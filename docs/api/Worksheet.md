[Microsoft Graph SDK](README.md) / Worksheet

# Worksheet

Worksheet models.

## Type Aliases

### Worksheet

> **Worksheet** = `object` & [`WriteWorksheet`](#writeworksheet)

Defined in: [src/models/Worksheet.ts:15](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Worksheet.ts#L15)

Represents a worksheet in a workbook.

#### Type declaration

##### id

> **id**: `string`

##### state

> **state**: `"visible"` \| `"hidden"` \| `"veryHidden"`

***

### WriteWorksheet

> **WriteWorksheet** = `object`

Defined in: [src/models/Worksheet.ts:25](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Worksheet.ts#L25)

Represents a worksheet to be written.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="name"></a> `name` | `string` | Name of the worksheet. | [src/models/Worksheet.ts:26](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Worksheet.ts#L26) |
| <a id="rows"></a> `rows` | [`Iterable`](https://www.typescriptlang.org/docs/handbook/iterators-and-generators.html#iterable-interface)\<[`WriteRow`](Row.md#writerow)\> \| `AsyncIterable`\<[`WriteRow`](Row.md#writerow)\> | Rows to write. | [src/models/Worksheet.ts:27](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Worksheet.ts#L27) |
