[Microsoft Graph SDK](README.md) / saveWorkbookAs

# saveWorkbookAs

Write workbook to Microsoft Sharepoint to a specific path.

## Functions

### saveWorkbookAs()

> **saveWorkbookAs**(`workbook`, `parentRef`, `path`, `options?`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`DriveItem` & `SiteRef` & `object` & `object`\>

Defined in: [src/tasks/saveWorkbookAs.ts:28](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/saveWorkbookAs.ts#L28)

Writes a workbook file to Microsoft SharePoint at a given location.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `workbook` | [`Workbook`](Handle.md#workbook) | Reference to the locally opened workbook. |
| `parentRef` | `DriveRef` \| `DriveItemRef` | Reference to the parent Drive or DriveItem where the file will be written. |
| `path` | `DriveItemPath` | Path where the workbook will be written in SharePoint. |
| `options?` | [`WriteOptions`](Options.md#writeoptions) | Options for writing, such as progress callback. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`DriveItem` & `SiteRef` & `object` & `object`\>

Resolves when the workbook has been written.
