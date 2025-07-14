[Microsoft Graph SDK](README.md) / readWorkbookByPath

# readWorkbookByPath

Reading a workbook from SharePoint by path.

## Functions

### readWorkbookByPath()

> **readWorkbookByPath**(`parentRef`, `itemPath`, `options`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Handle`](Handle.md#handle)\>

Defined in: [src/tasks/readWorkbookByPath.ts:25](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/readWorkbookByPath.ts#L25)

Reads a workbook file from a SharePoint drive by its path, supporting wildcards in the filename.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `parentRef` | `DriveRef` \| `DriveItemRef` | Reference to the parent drive or folder. |
| `itemPath` | `DriveItemPath` | Path to the file, may include wildcards in the filename. |
| `options` | [`ReadOptions`](ReadOptions.md#readoptions) | - |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Handle`](Handle.md#handle)\>

Reference to the locally opened workbook.

#### Throws

If the file path is invalid or no matching file is found.
