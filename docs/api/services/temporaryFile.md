[Microsoft Graph SDK](../README.md) / services/temporaryFile

# services/temporaryFile

## Functions

### getTemporaryFilePath()

> **getTemporaryFilePath**(`extension`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`LocalFilePath`](../LocalFilePath.md#localfilepath)\>

Defined in: [src/services/temporaryFile.ts:7](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/temporaryFile.ts#L7)

#### Parameters

| Parameter | Type | Default value |
| ------ | ------ | ------ |
| `extension` | `string` | `".tmp"` |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`LocalFilePath`](../LocalFilePath.md#localfilepath)\>

***

### withTemporaryFile()

> **withTemporaryFile**(`extension`, `context`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`void`\>

Defined in: [src/services/temporaryFile.ts:15](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/temporaryFile.ts#L15)

#### Parameters

| Parameter | Type |
| ------ | ------ |
| `extension` | `string` |
| `context` | (`file`) => [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`void`\> |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`void`\>
