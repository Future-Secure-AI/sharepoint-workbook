[Microsoft Graph SDK](README.md) / saveWorkbookAs

# saveWorkbookAs

Write workbook to Microsoft Sharepoint to a specific path.

## Type Aliases

### SaveWorkbookOptions

> **SaveWorkbookOptions** = `object`

Defined in: [src/tasks/saveWorkbookAs.ts:25](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/saveWorkbookAs.ts#L25)

Options for writing a workbook file.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="ifexists"></a> `ifExists?` | `"fail"` \| `"replace"` \| `"rename"` | Behavior if the file already exists. | [src/tasks/saveWorkbookAs.ts:26](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/saveWorkbookAs.ts#L26) |
| <a id="maxchunksize"></a> `maxChunkSize?` | `number` | Maximum chunk size in bytes for writing. | [src/tasks/saveWorkbookAs.ts:28](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/saveWorkbookAs.ts#L28) |
| <a id="progress"></a> `progress?` | (`bytes`) => `void` | Progress callback, receives bytes processed. | [src/tasks/saveWorkbookAs.ts:27](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/saveWorkbookAs.ts#L27) |

## Functions

### saveWorkbookAs()

> **saveWorkbookAs**(`workbook`, `parentRef`, `path`, `options?`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`DriveItem` & `SiteRef` & `object` & `object`\>

Defined in: [src/tasks/saveWorkbookAs.ts:44](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/saveWorkbookAs.ts#L44)

Writes a workbook file to Microsoft SharePoint at a given location.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `workbook` | [`Workbook`](Handle.md#workbook) | Reference to the locally opened workbook. |
| `parentRef` | `DriveRef` \| `DriveItemRef` | Reference to the parent Drive or DriveItem where the file will be written. |
| `path` | `DriveItemPath` | Path where the workbook will be written in SharePoint. |
| `options?` | [`SaveWorkbookOptions`](#saveworkbookoptions) | Options for writing, such as progress callback. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`DriveItem` & `SiteRef` & `object` & `object`\>

Resolves when the workbook has been written.

#### Remarks

See https://docs.aspose.com/cells/cpp/supported-file-formats/ for supported file formats. It cannot exceed SharePoint's file size limit of 250GB.
For size indication, a particular 700MB CSV file compresses down to about:
 - ~100MB XLSX
 - ~30MB XLSB
 - ~12MB XLS
