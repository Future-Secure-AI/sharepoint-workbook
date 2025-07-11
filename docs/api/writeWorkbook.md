[Microsoft Graph SDK](README.md) / writeWorkbook

# writeWorkbook

Write a workbook.

## Type Aliases

### WriteOptions

> **WriteOptions** = `object`

Defined in: [src/tasks/writeWorkbook.ts:51](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L51)

Options for writing a workbook file.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="compressionlevel"></a> `compressionLevel?` | `number` | Compression level for the output .xlsx zip file (0-9, default 6) | [src/tasks/writeWorkbook.ts:56](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L56) |
| <a id="ifalreadyexists"></a> `ifAlreadyExists?` | `"fail"` \| `"replace"` \| `"rename"` | What to do if the file already exists. | [src/tasks/writeWorkbook.ts:52](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L52) |
| <a id="maxchunksize"></a> `maxChunkSize?` | `number` | Maximum chunk size for upload (in bytes). | [src/tasks/writeWorkbook.ts:53](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L53) |
| <a id="progress"></a> `progress?` | (`update`) => `void` | Progress callback. | [src/tasks/writeWorkbook.ts:54](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L54) |
| <a id="workingfolder"></a> `workingFolder?` | `string` | Working folder for temporary file storage. Defaults to the `WORKING_FOLDER` env, then the OS temporary folder if not set. | [src/tasks/writeWorkbook.ts:55](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L55) |

***

### WriteProgress

> **WriteProgress** = `object`

Defined in: [src/tasks/writeWorkbook.ts:34](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L34)

Progress information for workbook writing operations.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="compressionratio"></a> `compressionRatio` | `number` | Ratio of compressed file size to original file size (0 to 1, where 1 is no compression) | [src/tasks/writeWorkbook.ts:37](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L37) |
| <a id="prepared"></a> `prepared` | `number` | Number of cells prepared for writing | [src/tasks/writeWorkbook.ts:35](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L35) |
| <a id="preparedpersecond"></a> `preparedPerSecond` | `number` | Number of cells prepared per second | [src/tasks/writeWorkbook.ts:38](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L38) |
| <a id="written"></a> `written` | `number` | Number of cells written to the destination | [src/tasks/writeWorkbook.ts:36](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L36) |
| <a id="writtenpersecond"></a> `writtenPerSecond` | `number` | Number of cells written per second | [src/tasks/writeWorkbook.ts:39](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L39) |

## Functions

### writeWorkbook()

> **writeWorkbook**(`parentRef`, `itemPath`, `sheets`, `options?`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`DriveItem` & `SiteRef` & `object` & `object`\>

Defined in: [src/tasks/writeWorkbook.ts:69](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L69)

**`Experimental`**

Writes a workbook (.xlsx) in the specified parent location with the provided rows for multiple sheets.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `parentRef` | `DriveRef` \| `DriveItemRef` | Reference to the parent drive or item where the file will be written. |
| `itemPath` | `DriveItemPath` | Path (including filename and extension) for the new workbook. |
| `sheets` | [`Record`](https://www.typescriptlang.org/docs/handbook/utility-types.html#recordkeys-type)\<`WorkbookWorksheetName`, [`Iterable`](https://www.typescriptlang.org/docs/handbook/iterators-and-generators.html#iterable-interface)\<[`Partial`](https://www.typescriptlang.org/docs/handbook/utility-types.html#partialtype)\<`Cell`\>[]\> \| `AsyncIterable`\<[`Partial`](https://www.typescriptlang.org/docs/handbook/utility-types.html#partialtype)\<`Cell`\>[]\>\> | Object where each key is a sheet name (WorkbookWorksheetName) and the value is an iterable or async iterable of row arrays. |
| `options?` | [`WriteOptions`](#writeoptions) | Options for conflict resolution, etc. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`DriveItem` & `SiteRef` & `object` & `object`\>

Written DriveItem with reference.

#### Throws

If the file extension is not supported.
