[Microsoft Graph SDK](README.md) / openWorkbookByPath

# openWorkbookByPath

Reading a workbook from SharePoint by path.

## Type Aliases

### OpenWorkbookOptions

> **OpenWorkbookOptions** = `object`

Defined in: [src/tasks/openWorkbook.ts:30](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/openWorkbook.ts#L30)

Options for reading a workbook file.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="progress"></a> `progress?` | (`bytes`) => `void` | Progress callback, receives bytes processed. | [src/tasks/openWorkbook.ts:31](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/openWorkbook.ts#L31) |

## Functions

### openWorkbook()

> **openWorkbook**(`parentRef`, `itemPath`, `options`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Workbook`](Handle.md#workbook)\>

Defined in: [src/tasks/openWorkbook.ts:46](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/openWorkbook.ts#L46)

Reads a workbook file from a SharePoint drive by its path, supporting wildcards in the filename.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `parentRef` | `DriveRef` \| `DriveItemRef` | Reference to the parent drive or folder. |
| `itemPath` | `DriveItemPath` | Path to the file, may include wildcards in the filename. |
| `options` | [`OpenWorkbookOptions`](#openworkbookoptions) | - |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Workbook`](Handle.md#workbook)\>

Reference to the locally opened workbook.

#### Throws

If the file path is invalid or no matching file is found.

#### Remarks

Supported workbooks are:
- Supported file type https://docs.aspose.com/cells/cpp/supported-file-formats/
- No more than 250GB
- No more than 1/4 of the memory available to Node (increase physical memory and `--max-old-space-size` if needed)
- Optionally compressed with GZip (with an appended .gz extension)
