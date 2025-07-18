/**
 * Cells and its properties in a worksheet.
 * @module Cell
 * @category Models
 */

/**
 * Represents the value of a cell. Can be a string, number, boolean, or Date.
 * @typedef {string | number | boolean | Date} CellValue
 */
export type CellValue = string | number | boolean | Date;

/**
 * Represents a cell in a worksheet, including its value, formatting, alignment, borders, and other properties.
 * @typedef {Object} Cell
 * @property {CellValue} value The value contained in the cell.
 * @property {string} formula The formula for the cell, empty if none.
 * @property {string} fontName Font family name for the cell text.
 * @property {number} fontSize Font size for the cell text.
 * @property {boolean} fontBold Whether the cell text is bold.
 * @property {boolean} fontItalic Whether the cell text is italic.
 * @property {string} fontColor Font color for the cell text (6-character hex color string, no hash).
 * @property {string} backgroundColor Background color of the cell (6-character hex color string, no hash).
 * @property {CellHorizontalAlignment} horizontalAlignment Horizontal alignment of the cell content.
 * @property {CellVerticalAlignment} verticalAlignment Vertical alignment of the cell content.
 * @property {number} rotationAngle Text rotation angle in degrees (0-180, or -90 for vertical text).
 * @property {boolean} isTextWrapped Whether text wrapping is enabled for the cell.
 * @property {CellBorderStyle} borderTopStyle Border style for the top edge of the cell.
 * @property {string} borderTopColor Border color for the top edge of the cell (6-character hex color string, no hash).
 * @property {CellBorderStyle} borderBottomStyle Border style for the bottom edge of the cell.
 * @property {string} borderBottomColor Border color for the bottom edge of the cell (6-character hex color string, no hash).
 * @property {CellBorderStyle} borderLeftStyle Border style for the left edge of the cell.
 * @property {string} borderLeftColor Border color for the left edge of the cell (6-character hex color string, no hash).
 * @property {CellBorderStyle} borderRightStyle Border style for the right edge of the cell.
 * @property {string} borderRightColor Border color for the right edge of the cell (6-character hex color string, no hash).
 * @property {string} numberFormat Number format string for the cell (e.g., Excel format).
 * @property {boolean} isLocked Whether the cell is locked (protected).
 * @property {number} indentLevel Indentation level for the cell content.
 * @property {boolean} shrinkToFit Whether to shrink text to fit the cell.
 * @property {CellMerge} merge Merge state of the cell.
 * @property {string} comment Comment or note attached to the cell.
 */
export type Cell = {
	value: CellValue;
	formula: string;
	fontName: string;
	fontSize: number;
	fontBold: boolean;
	fontItalic: boolean;
	fontColor: string;
	backgroundColor: string;
	horizontalAlignment: CellHorizontalAlignment;
	verticalAlignment: CellVerticalAlignment;
	rotationAngle: number;
	isTextWrapped: boolean;
	borderTopStyle: CellBorderStyle;
	borderTopColor: string;
	borderBottomStyle: CellBorderStyle;
	borderBottomColor: string;
	borderLeftStyle: CellBorderStyle;
	borderLeftColor: string;
	borderRightStyle: CellBorderStyle;
	borderRightColor: string;
	numberFormat: string;
	isLocked: boolean;
	indentLevel: number;
	shrinkToFit: boolean;
	merge: CellMerge;
	comment: string;
};

/**
 * Indicates if the cell is merged, and in what direction.
 *
 * - "up": merged with the cell above
 * - "left": merged with the cell to the left
 * - "up-left": merged with the cell above and to the left
 * - null: not merged
 *
 * @typedef {('up'|'left'|'up-left'|null)} CellMerge
 */
export type CellMerge = "up" | "left" | "up-left" | null;

/**
 * Border style for a cell edge.
 * @typedef {('thin'|'medium'|'thick'|'dashed'|'dotted'|'double')} CellBorderStyle
 */
export type CellBorderStyle = "thin" | "medium" | "thick" | "dashed" | "dotted" | "double";

/**
 * Horizontal alignment options for cell content.
 * @typedef {('left'|'center'|'right')} CellHorizontalAlignment
 */
export type CellHorizontalAlignment = "left" | "center" | "right";

/**
 * Vertical alignment options for cell content.
 * @typedef {('top'|'middle'|'bottom')} CellVerticalAlignment
 */
export type CellVerticalAlignment = "top" | "middle" | "bottom";
