[Microsoft Graph SDK](README.md) / createWorkbook

# createWorkbook

Copy a drive item.

## Type Aliases

### CreateOptions

> **CreateOptions** = `object`

Defined in: [src/tasks/createWorkbook.ts:26](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/createWorkbook.ts#L26)

Options for creating a new workbook file.

#### Properties

| Property | Type | Defined in |
| ------ | ------ | ------ |
| <a id="conflictbehavior"></a> `conflictBehavior?` | `"fail"` \| `"replace"` \| `"rename"` | [src/tasks/createWorkbook.ts:27](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/createWorkbook.ts#L27) |
| <a id="maxchunksize"></a> `maxChunkSize?` | `number` | [src/tasks/createWorkbook.ts:28](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/createWorkbook.ts#L28) |
| <a id="progress"></a> `progress?` | (`preparedCount`, `writtenCount`, `preparedPerSecond`, `writtenPerSecond`) => `void` | [src/tasks/createWorkbook.ts:29](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/createWorkbook.ts#L29) |

## Functions

### createWorkbook()

> **createWorkbook**(`parentRef`, `itemPath`, `sheets`, `options?`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`DriveItem` & `SiteRef` & `object` & `object`\>

Defined in: [src/tasks/createWorkbook.ts:41](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/createWorkbook.ts#L41)

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
