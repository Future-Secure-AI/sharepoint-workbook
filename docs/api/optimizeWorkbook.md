[Microsoft Graph SDK](README.md) / optimizeWorkbook

# optimizeWorkbook

Optimize a opened workbook file by recompressing with a specified compression level.

## Functions

### optimizeWorkbook()

> **optimizeWorkbook**(`hdl`, `options`): [`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`number`\>

Defined in: [src/tasks/optimizeWorkbook.ts:22](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/optimizeWorkbook.ts#L22)

Optimizes an opened workbook by recompressing it.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `hdl` | [`Handle`](Handle.md#handle) | Reference to the opened workbook. |
| `options` | [`OptimizeOptions`](OptimizeOptions.md#optimizeoptions) | Options for optimization, including compression level. |

#### Returns

[`Promise`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise)\<`number`\>

The ratio of the output file size to the input file size.

#### Throws

If the optimization fails.
