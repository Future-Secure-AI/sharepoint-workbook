[Microsoft Graph SDK](README.md) / findWorksheet

# findWorksheet

Find a worksheet in a workbook by name.

## Functions

### findWorksheet()

> **findWorksheet**(`workbook`, `search`): `Worksheet`

Defined in: [src/tasks/findWorksheet.ts:18](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/tasks/findWorksheet.ts#L18)

Finds the first worksheet in the workbook whose name matches the given string or glob pattern (case-insensitive).

#### Parameters

| Parameter | Type | Description |
| ------ | ------ | ------ |
| `workbook` | [`Workbook`](Handle.md#workbook) | The workbook to search. |
| `search` | `string` | The worksheet name or glob pattern to match. |

#### Returns

`Worksheet`

The first matching worksheet.

#### Throws

If no worksheet matches the pattern.
