[Microsoft Graph SDK](../README.md) / services/workingFolder

# services/workingFolder

## Functions

### createHandleId()

> **createHandleId**(): [`HandleId`](../Handle.md#handleid-1)

Defined in: [src/services/workingFolder.ts:20](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/workingFolder.ts#L20)

#### Returns

[`HandleId`](../Handle.md#handleid-1)

***

### getLatestRevisionFilePath()

> **getLatestRevisionFilePath**(`id`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`LocalFilePath`](../LocalFilePath.md#localfilepath)\>

Defined in: [src/services/workingFolder.ts:24](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/workingFolder.ts#L24)

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `id` | [`HandleId`](../Handle.md#handleid-1) |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`LocalFilePath`](../LocalFilePath.md#localfilepath)\>

***

### getNextRevisionFilePath()

> **getNextRevisionFilePath**(`id`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`LocalFilePath`](../LocalFilePath.md#localfilepath)\>

Defined in: [src/services/workingFolder.ts:42](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/workingFolder.ts#L42)

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `id` | [`HandleId`](../Handle.md#handleid-1) |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`LocalFilePath`](../LocalFilePath.md#localfilepath)\>
