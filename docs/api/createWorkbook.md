[Microsoft Graph SDK](README.md) / createWorkbook

# createWorkbook

Create a workbook.

## Type Aliases

### CreateOptions

> **CreateOptions** = `object`

Defined in: [src/tasks/writeWorkbook.ts:30](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L30)

Options for creating a new workbook file.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="ifalreadyexists"></a> `ifAlreadyExists?` | `"fail"` \| `"replace"` \| `"rename"` | How to resolve if the file already exists. | [src/tasks/writeWorkbook.ts:31](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L31) |
| <a id="maxchunksize"></a> `maxChunkSize?` | `number` | Maximum chunk size for upload (in bytes). | [src/tasks/writeWorkbook.ts:32](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L32) |
| <a id="progress"></a> `progress?` | (`preparedCount`, `writtenCount`, `preparedPerSecond`, `writtenPerSecond`) => `void` | Progress callback. | [src/tasks/writeWorkbook.ts:33](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L33) |
| <a id="workingfolder"></a> `workingFolder?` | `string` | Working folder for temporary file storage. Defaults to the `WORKING_FOLDER` env, then the OS temporary folder if not set. | [src/tasks/writeWorkbook.ts:34](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L34) |

## Functions

### writeWorkbook()

> **writeWorkbook**(`parentRef`, `itemPath`, `sheets`, `options?`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`DriveItem` & `SiteRef` & `object` & `object`\>

Defined in: [src/tasks/writeWorkbook.ts:47](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/writeWorkbook.ts#L47)

**`Experimental`**

Creates a new workbook (.xlsx) in the specified parent location with the provided rows for multiple sheets.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `parentRef` | `DriveRef` \| `DriveItemRef` | Reference to the parent drive or item where the file will be created. |
| `itemPath` | `DriveItemPath` | Path (including filename and extension) for the new workbook. |
| `sheets` | [`Record`](https://www.typescriptlang.org/docs/handbook/utility-types.html#recordkeys-type)\<`WorkbookWorksheetName`, [`Iterable`](https://www.typescriptlang.org/docs/handbook/iterators-and-generators.html#iterable-interface)\<[`Partial`](https://www.typescriptlang.org/docs/handbook/utility-types.html#partialtype)\<`Cell`\>[]\> \| `AsyncIterable`\<[`Partial`](https://www.typescriptlang.org/docs/handbook/utility-types.html#partialtype)\<`Cell`\>[]\>\> | Object where each key is a sheet name (WorkbookWorksheetName) and the value is an iterable or async iterable of row arrays. |
| `options?` | [`CreateOptions`](#createoptions) | Options for conflict resolution, etc. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`DriveItem` & `SiteRef` & `object` & `object`\>

Created DriveItem with reference.

#### Throws

If the file extension is not supported.
