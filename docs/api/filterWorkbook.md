[Microsoft Graph SDK](README.md) / filterWorkbook

# filterWorkbook

Filter out unwanted rows and columns from a workbook.

## Type Aliases

### Filter

> **Filter** = `object`

Defined in: [src/tasks/filterWorkbook.ts:15](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/filterWorkbook.ts#L15)

Filter options for filtering a workbook.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="columnfilter"></a> `columnFilter?` | (`header`, `index`) => `boolean` | Function to determine if a column should be included, based on header and index. Return true to include the column, or false to omit it. | [src/tasks/filterWorkbook.ts:22](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/filterWorkbook.ts#L22) |
| <a id="progress"></a> `progress?` | (`rows`) => `void` | - | [src/tasks/filterWorkbook.ts:29](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/filterWorkbook.ts#L29) |
| <a id="rowfilter"></a> `rowFilter?` | (`cells`) => `boolean` | Function to determine if a row should be included, based on cell values. Return true to include the row, or false to omit it. | [src/tasks/filterWorkbook.ts:27](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/filterWorkbook.ts#L27) |
| <a id="skiprows"></a> `skipRows?` | `number` | Number of rows to skip from the top (e.g., header rows). | [src/tasks/filterWorkbook.ts:17](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/filterWorkbook.ts#L17) |

## Functions

### filterWorkbook()

> **filterWorkbook**(`workbook`, `filter`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`void`\>

Defined in: [src/tasks/filterWorkbook.ts:40](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/filterWorkbook.ts#L40)

Filter out unwanted rows and columns from a workbook. All styling is lost when filtering.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `workbook` | [`Workbook`](Handle.md#workbook) | Workbook handle to filter. |
| `filter` | [`Filter`](#filter) | Filter options to apply (skipRows, column, row). |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`void`\>

A promise that resolves when the filtering is complete.
