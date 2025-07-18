[Microsoft Graph SDK](README.md) / Options

# Options

Options for operation.

## Type Aliases

### ReadOptions

> **ReadOptions** = `object`

Defined in: [src/models/Options.ts:12](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Options.ts#L12)

Options for reading a workbook file.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="progress"></a> `progress?` | (`bytes`) => `void` | Progress callback, receives bytes processed. | [src/models/Options.ts:13](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Options.ts#L13) |

***

### WriteOptions

> **WriteOptions** = `object`

Defined in: [src/models/Options.ts:22](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Options.ts#L22)

Options for writing a workbook file.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="ifexists"></a> `ifExists?` | `"fail"` \| `"replace"` \| `"rename"` | Behavior if the file already exists. | [src/models/Options.ts:23](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Options.ts#L23) |
| <a id="maxchunksize"></a> `maxChunkSize?` | `number` | Maximum chunk size in bytes for writing. | [src/models/Options.ts:25](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Options.ts#L25) |
| <a id="progress-1"></a> `progress?` | (`bytes`) => `void` | Progress callback, receives bytes processed. | [src/models/Options.ts:24](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Options.ts#L24) |
