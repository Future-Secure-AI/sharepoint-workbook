[Microsoft Graph SDK](README.md) / ReadOptions

# ReadOptions

Configuration for a read operation.

## Type Aliases

### ReadOptions

> **ReadOptions** = `object`

Defined in: [src/models/ReadOptions.ts:13](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/ReadOptions.ts#L13)

Options for reading a workbook file.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="defaultworksheetname"></a> `defaultWorksheetName?` | `WorkbookWorksheetName` | Default worksheet name to use when importing a CSV file. | [src/models/ReadOptions.ts:14](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/ReadOptions.ts#L14) |
| <a id="progress"></a> `progress?` | (`bytes`) => `void` | Progress callback, receives bytes processed. | [src/models/ReadOptions.ts:15](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/ReadOptions.ts#L15) |
