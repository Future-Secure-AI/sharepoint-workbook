[Microsoft Graph SDK](../README.md) / models/Cell

# models/Cell

## Type Aliases

### Cell

> **Cell** = `object`

Defined in: [src/models/Cell.ts:9](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L9)

Represents a cell in a worksheet, including its value, formatting, alignment, borders, and other properties.

#### Properties

| Property | Type | Description | Defined in |
| ------ | ------ | ------ | ------ |
| <a id="backgroundcolor"></a> `backgroundColor` | `string` | Background color of the cell (6-character hex color string, no hash). | [src/models/Cell.ts:27](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L27) |
| <a id="borderbottomcolor"></a> `borderBottomColor` | `string` | Border color for the bottom edge of the cell. (6-character hex color string, no hash) | [src/models/Cell.ts:45](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L45) |
| <a id="borderbottomstyle"></a> `borderBottomStyle` | [`CellBorderStyle`](#cellborderstyle) | Border style for the bottom edge of the cell. | [src/models/Cell.ts:43](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L43) |
| <a id="borderleftcolor"></a> `borderLeftColor` | `string` | Border color for the left edge of the cell. (6-character hex color string, no hash) | [src/models/Cell.ts:49](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L49) |
| <a id="borderleftstyle"></a> `borderLeftStyle` | [`CellBorderStyle`](#cellborderstyle) | Border style for the left edge of the cell. | [src/models/Cell.ts:47](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L47) |
| <a id="borderrightcolor"></a> `borderRightColor` | `string` | Border color for the right edge of the cell. (6-character hex color string, no hash) | [src/models/Cell.ts:53](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L53) |
| <a id="borderrightstyle"></a> `borderRightStyle` | [`CellBorderStyle`](#cellborderstyle) | Border style for the right edge of the cell. | [src/models/Cell.ts:51](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L51) |
| <a id="bordertopcolor"></a> `borderTopColor` | `string` | Border color for the top edge of the cell.(6-character hex color string, no hash) | [src/models/Cell.ts:41](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L41) |
| <a id="bordertopstyle"></a> `borderTopStyle` | [`CellBorderStyle`](#cellborderstyle) | Border style for the top edge of the cell. | [src/models/Cell.ts:39](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L39) |
| <a id="comment"></a> `comment` | `string` | Comment or note attached to the cell. | [src/models/Cell.ts:71](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L71) |
| <a id="fontbold"></a> `fontBold` | `boolean` | Whether the cell text is bold. | [src/models/Cell.ts:20](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L20) |
| <a id="fontcolor"></a> `fontColor` | `string` | Font color for the cell text (6-character hex color string, no hash). | [src/models/Cell.ts:24](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L24) |
| <a id="fontitalic"></a> `fontItalic` | `boolean` | Whether the cell text is italic. | [src/models/Cell.ts:22](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L22) |
| <a id="fontname"></a> `fontName` | `string` | Font family name for the cell text. | [src/models/Cell.ts:16](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L16) |
| <a id="fontsize"></a> `fontSize` | `number` | Font size for the cell text. | [src/models/Cell.ts:18](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L18) |
| <a id="formula"></a> `formula` | `string` | The formula for the cell, empty if none. | [src/models/Cell.ts:13](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L13) |
| <a id="horizontalalignment"></a> `horizontalAlignment` | [`CellHorizontalAlignment`](#cellhorizontalalignment-1) | Horizontal alignment of the cell content. | [src/models/Cell.ts:30](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L30) |
| <a id="indentlevel"></a> `indentLevel` | `number` | Indentation level for the cell content. | [src/models/Cell.ts:62](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L62) |
| <a id="islocked"></a> `isLocked` | `boolean` | Whether the cell is locked (protected). | [src/models/Cell.ts:59](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L59) |
| <a id="istextwrapped"></a> `isTextWrapped` | `boolean` | Whether text wrapping is enabled for the cell. | [src/models/Cell.ts:36](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L36) |
| <a id="merge"></a> `merge` | [`CellMerge`](#cellmerge-1) | Merge state of the cell. | [src/models/Cell.ts:68](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L68) |
| <a id="numberformat"></a> `numberFormat` | `string` | Number format string for the cell (e.g., Excel format). | [src/models/Cell.ts:56](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L56) |
| <a id="rotationangle"></a> `rotationAngle` | `number` | Text rotation angle in degrees (0-180, or -90 for vertical text). | [src/models/Cell.ts:34](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L34) |
| <a id="shrinktofit"></a> `shrinkToFit` | `boolean` | Whether to shrink text to fit the cell. | [src/models/Cell.ts:65](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L65) |
| <a id="value"></a> `value` | [`CellValue`](#cellvalue-1) | The value contained in the cell. | [src/models/Cell.ts:11](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L11) |
| <a id="verticalalignment"></a> `verticalAlignment` | [`CellVerticalAlignment`](#cellverticalalignment-1) | Vertical alignment of the cell content. | [src/models/Cell.ts:32](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L32) |

***

### CellBorderStyle

> **CellBorderStyle** = `"thin"` \| `"medium"` \| `"thick"` \| `"dashed"` \| `"dotted"` \| `"double"`

Defined in: [src/models/Cell.ts:86](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L86)

Border style for a cell edge.

***

### CellHorizontalAlignment

> **CellHorizontalAlignment** = `"left"` \| `"center"` \| `"right"`

Defined in: [src/models/Cell.ts:91](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L91)

Horizontal alignment options for cell content.

***

### CellMerge

> **CellMerge** = `"up"` \| `"left"` \| `"up-left"` \| `null`

Defined in: [src/models/Cell.ts:81](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L81)

Indicates if the cell is merged, and in what direction.
- "up": merged with the cell above
- "left": merged with the cell to the left
- "up-left": merged with the cell above and to the left
- null: not merged

***

### CellValue

> **CellValue** = `string` \| `number` \| `boolean` \| [`Date`](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Date)

Defined in: [src/models/Cell.ts:4](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L4)

Represents the value of a cell. Can be a string, number, boolean, or Date.

***

### CellVerticalAlignment

> **CellVerticalAlignment** = `"top"` \| `"middle"` \| `"bottom"`

Defined in: [src/models/Cell.ts:96](https://github.com/Future-Secure-AI/sharepoint-workbook/blob/main/src/models/Cell.ts#L96)

Vertical alignment options for cell content.
