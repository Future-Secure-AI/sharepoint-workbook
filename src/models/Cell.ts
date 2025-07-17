export type CellValue = string | number | boolean | Date;

export type Cell = {
	value: CellValue;
	formula: string;
	style: CellStyle;
	merge: CellMerge;
	comment: string;
};

export type CellStyle = {
	font: CellFont;
	backgroundColor: string;
	horizontalAlignment: "left" | "center" | "right";
	verticalAlignment: "top" | "middle" | "bottom";
	borders: CellBorders;
	numberFormat: string;
	locked: boolean;
	wrapText: boolean;
};

export type CellBorder = {
	style: "thin" | "medium" | "thick" | "dashed" | "dotted" | "double";
	color: string;
};

export type CellBorders = {
	top: CellBorder;
	bottom: CellBorder;
	left: CellBorder;
	right: CellBorder;
};

export type CellFont = {
	name: string;
	size: number;
	bold: boolean;
	italic: boolean;
	color: string;
};

/**
 * If the cell is merged, and in what direction.
 * "up" means the cell is merged with the cell above it,
 * "left" means the cell is merged with the cell to the left,
 * "up-left" means the cell is merged with the cell above and to the left,
 * and null means the cell is not merged.
 */
export type CellMerge = "up" | "left" | "up-left" | null;
