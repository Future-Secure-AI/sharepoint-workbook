[Microsoft Graph SDK](README.md) / getWorksheet

# getWorksheet

Get a worksheet from a workbook by its exact name.

## Functions

### getWorksheet()

> **getWorksheet**(`workbook`, `name`): `Worksheet`

Defined in: [src/tasks/getWorksheet.ts:18](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/getWorksheet.ts#L18)

Returns the worksheet with the given name from the workbook.

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `workbook` | [`Workbook`](Handle.md#workbook) | The workbook to search. |
| `name` | [`WorksheetName`](Worksheet.md#worksheetname) | The exact name of the worksheet to retrieve. |

#### Returns

`Worksheet`

The worksheet with the specified name.

#### Throws

If the worksheet is not found.
