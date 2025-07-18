[Microsoft Graph SDK](README.md) / filterWorkbookRows

# filterWorkbookRows

Filter out unwanted rows from a workbook.

## Type Aliases

### FilterWorkbookRowsOptions

> **FilterWorkbookRowsOptions** = `object`

Defined in: src/tasks/filterWorkbookRows.ts:15

Options for filtering workbook rows.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="progress"></a> `progress?` | (`rows`) => `void` | Optional callback to report progress, called with the number of processed rows. | src/tasks/filterWorkbookRows.ts:17 |
| <a id="skiprows"></a> `skipRows?` | `number` | Number of rows to skip from the top of each worksheet before filtering. | src/tasks/filterWorkbookRows.ts:16 |

## Functions

### filterWorkbookRows()

> **filterWorkbookRows**(`workbook`, `rowFilter`, `options`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`void`\>

Defined in: src/tasks/filterWorkbookRows.ts:27

Filter out unwanted rows from a workbook.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `workbook` | [`Workbook`](Handle.md#workbook) | Workbook handle to filter. |
| `rowFilter` | (`cells`) => `boolean` | Function to determine if a row should be included, based on cell values. Return true to include the row, or false to omit it. |
| `options` | [`FilterWorkbookRowsOptions`](#filterworkbookrowsoptions) | Row filter options to apply (skipRows, progress). |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`void`\>

A promise that resolves when the filtering is complete.
