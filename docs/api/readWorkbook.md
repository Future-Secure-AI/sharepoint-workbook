[Microsoft Graph SDK](README.md) / readWorkbook

# readWorkbook

Read workbook from Microsoft SharePoint.

## Functions

### readWorkbook()

> **readWorkbook**(`itemRef`, `options?`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Handle`](Handle.md#handle)\>

Defined in: [src/tasks/readWorkbook.ts:29](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/readWorkbook.ts#L29)

Reads a workbook file (.xlsx or .csv) from a Microsoft Graph.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `itemRef` | `SiteRef` & `object` & `object` & [`Partial`](https://www.typescriptlang.org/docs/handbook/utility-types.html#partialtype)\<`DriveItem`\> | Reference to the DriveItem to read from. |
| `options?` | [`ReadOptions`](ReadOptions.md#readoptions) | Options for reading, such as default worksheet name for CSV. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`Handle`](Handle.md#handle)\>

Reference to the locally opened workbook.

#### Throws

If the file extension is not supported.
