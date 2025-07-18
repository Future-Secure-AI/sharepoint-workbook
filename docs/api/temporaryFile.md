[Microsoft Graph SDK](README.md) / temporaryFile

# temporaryFile

Utilities for creating temporary file paths for workbook operations.

## Functions

### getTemporaryFilePath()

> **getTemporaryFilePath**(`extension?`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`LocalFilePath`](LocalFilePath.md#localfilepath)\>

Defined in: [src/services/temporaryFile.ts:18](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/services/temporaryFile.ts#L18)

Generate a unique temporary file path, ensuring the directory exists.

#### Parameters

| Parameter | Type | Default value | Description |
| ------ | ------ | ------ | ------ |
| `extension?` | `string` | `".tmp"` | The file extension to use for the temporary file. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<[`LocalFilePath`](LocalFilePath.md#localfilepath)\>

The generated temporary file path.

#### Remarks

The file is not created, only the path is generated. The root folder is determined by the `WORKING_FOLDER` environment variable if set, otherwise defaults to a subdirectory in the OS temp directory.
