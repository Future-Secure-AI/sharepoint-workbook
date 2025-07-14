[Microsoft Graph SDK](README.md) / writeWorkbook

# writeWorkbook

Write a locally opened workbook back to Microsoft SharePoint.

## Functions

### writeWorkbook()

> **writeWorkbook**(`hdl`, `options?`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`void`\>

Defined in: [src/tasks/writeWorkbook.ts:24](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L24)

Write a locally opened workbook back to Microsoft SharePoint, overwriting the previous file.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `hdl` | [`Handle`](Handle.md#handle) | Reference to the locally opened workbook, must include an itemRef for overwrite. |
| `options?` | [`WriteOptions`](WriteOptions.md#writeoptions) | Options for writing, such as conflict behavior, chunk size, and progress callback. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`void`\>

Resolves when the upload is complete.

#### Throws

If the workbook cannot be overwritten or required metadata is missing.
