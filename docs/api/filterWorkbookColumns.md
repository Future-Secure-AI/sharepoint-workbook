[Microsoft Graph SDK](README.md) / filterWorkbookColumns

# filterWorkbookColumns

Filter out unwanted columns from a workbook.

## Functions

### filterWorkbookColumns()

> **filterWorkbookColumns**(`workbook`, `columnFilter`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`void`\>

Defined in: src/tasks/filterWorkbookColumns.ts:15

Filter out unwanted columns from a workbook.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `workbook` | `Workbook` | Workbook handle to filter. |
| `columnFilter` | (`header`, `index`) => `boolean` | Function to determine if a column should be included, based on header and index. Return true to include the column, or false to omit it. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`void`\>

A promise that resolves when the filtering is complete.
