[Microsoft Graph SDK](README.md) / Handle

# Handle

Reference to an opened workbook.

## Type Aliases

### Handle

> **Handle** = `object`

Defined in: [src/models/Handle.ts:13](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Handle.ts#L13)

A reference to an opened workbook.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="id"></a> `id` | [`HandleId`](#handleid-1) | Unique identifier for the handle. | [src/models/Handle.ts:14](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Handle.ts#L14) |
| <a id="itemref"></a> `itemRef?` | `DriveItemRef` | (Optional) Reference to the associated DriveItem in Microsoft Graph. | [src/models/Handle.ts:15](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Handle.ts#L15) |

***

### HandleId

> **HandleId** = `string` & `object`

Defined in: [src/models/Handle.ts:21](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Handle.ts#L21)

Unique Handle identifier.

#### Type declaration

##### \_\_brand

> `readonly` **\_\_brand**: unique `symbol`
