/**
 * Represents the value of a cell. Can be a string, number, boolean, or Date.
 */
export type CellValue = string | number | boolean | Date;

/**
 * Represents a cell in a worksheet, including its value, formatting, alignment, borders, and other properties.
 */
export type Cell = {
	/** The value contained in the cell. */
	value: CellValue;
	/** The formula for the cell, empty if none. */
	formula: string;

	/** Font family name for the cell text. */
	fontName: string;
	/** Font size for the cell text. */
	fontSize: number;
	/** Whether the cell text is bold. */
	fontBold: boolean;
	/** Whether the cell text is italic. */
	fontItalic: boolean;
	/** Font color for the cell text (6-character hex color string, no hash). */
	fontColor: string;

	/** Background color of the cell (6-character hex color string, no hash). */
	backgroundColor: string;

	/** Horizontal alignment of the cell content. */
	horizontalAlignment: CellHorizontalAlignment;
	/** Vertical alignment of the cell content. */
	verticalAlignment: CellVerticalAlignment;
	/** Text rotation angle in degrees (0-180, or -90 for vertical text). */
	rotationAngle: number;
	/** Whether text wrapping is enabled for the cell. */
	isTextWrapped: boolean;

	/** Border style for the top edge of the cell. */
	borderTopStyle: CellBorderStyle;
	/** Border color for the top edge of the cell.(6-character hex color string, no hash) */
	borderTopColor: string;
	/** Border style for the bottom edge of the cell. */
	borderBottomStyle: CellBorderStyle;
	/** Border color for the bottom edge of the cell. (6-character hex color string, no hash)*/
	borderBottomColor: string;
	/** Border style for the left edge of the cell. */
	borderLeftStyle: CellBorderStyle;
	/** Border color for the left edge of the cell. (6-character hex color string, no hash)*/
	borderLeftColor: string;
	/** Border style for the right edge of the cell. */
	borderRightStyle: CellBorderStyle;
	/** Border color for the right edge of the cell. (6-character hex color string, no hash)*/
	borderRightColor: string;

	/** Number format string for the cell (e.g., Excel format). */
	numberFormat: string;

	/** Whether the cell is locked (protected). */
	isLocked: boolean;

	/** Indentation level for the cell content. */
	indentLevel: number;

	/** Whether to shrink text to fit the cell. */
	shrinkToFit: boolean;

	/** Merge state of the cell. */
	merge: CellMerge;

	/** Comment or note attached to the cell. */
	comment: string;
};

/**
 * Indicates if the cell is merged, and in what direction.
 * - "up": merged with the cell above
 * - "left": merged with the cell to the left
 * - "up-left": merged with the cell above and to the left
 * - null: not merged
 */
export type CellMerge = "up" | "left" | "up-left" | null;

/**
 * Border style for a cell edge.
 */
export type CellBorderStyle = "thin" | "medium" | "thick" | "dashed" | "dotted" | "double";

/**
 * Horizontal alignment options for cell content.
 */
export type CellHorizontalAlignment = "left" | "center" | "right";

/**
 * Vertical alignment options for cell content.
 */
export type CellVerticalAlignment = "top" | "middle" | "bottom";
