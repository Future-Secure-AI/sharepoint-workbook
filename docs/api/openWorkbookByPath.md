[Microsoft Graph SDK](README.md) / openWorkbookByPath

# openWorkbookByPath

Reading a workbook from SharePoint by path.

## Functions

### openWorkbook()

> **openWorkbook**(`parentRef`, `itemPath`, `options`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Workbook`](Handle.md#workbook)\>

Defined in: [src/tasks/openWorkbook.ts:32](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/openWorkbook.ts#L32)

Reads a workbook file from a SharePoint drive by its path, supporting wildcards in the filename.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `parentRef` | `DriveRef` \| `DriveItemRef` | Reference to the parent drive or folder. |
| `itemPath` | `DriveItemPath` | Path to the file, may include wildcards in the filename. |
| `options` | [`ReadOptions`](Options.md#readoptions) | - |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Workbook`](Handle.md#workbook)\>

Reference to the locally opened workbook.

#### Throws

If the file path is invalid or no matching file is found.
