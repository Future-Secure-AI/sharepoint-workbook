[Microsoft Graph SDK](README.md) / WriteOptions

# WriteOptions

Options for writing a workbook file.

## Type Aliases

### WriteOptions

> **WriteOptions** = `object`

Defined in: [src/models/WriteOptions.ts:9](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/WriteOptions.ts#L9)

#### Properties

| Property | Type | Defined in |
| ------ | ------ | ------ |
| <a id="ifexists"></a> `ifExists?` | `"fail"` \| `"replace"` \| `"rename"` | [src/models/WriteOptions.ts:10](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/WriteOptions.ts#L10) |
| <a id="maxchunksize"></a> `maxChunkSize?` | `number` | [src/models/WriteOptions.ts:12](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/WriteOptions.ts#L12) |
| <a id="progress"></a> `progress?` | (`bytes`) => `void` | [src/models/WriteOptions.ts:11](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/WriteOptions.ts#L11) |
