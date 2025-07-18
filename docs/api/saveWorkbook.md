[Microsoft Graph SDK](README.md) / saveWorkbook

# saveWorkbook

Write opened workbook back to Microsoft SharePoint.

## Functions

### saveWorkbook()

> **saveWorkbook**(`handle`, `options?`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`SiteRef` & `object` & `object` & `DriveItem`\>

Defined in: [src/tasks/saveWorkbook.ts:22](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/saveWorkbook.ts#L22)

Write a locally opened workbook back to Microsoft SharePoint, overwriting the previous file.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `handle` | [`Workbook`](Handle.md#workbook) | Reference to the locally opened workbook, must include an itemRef for overwrite. |
| `options?` | [`SaveWorkbookOptions`](saveWorkbookAs.md#saveworkbookoptions) | Options for writing, such as conflict behavior, chunk size, and progress callback. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`SiteRef` & `object` & `object` & `DriveItem`\>

Resolves when the upload is complete.

#### Throws

If the workbook cannot be overwritten or required metadata is missing.
