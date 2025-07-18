/**
 * Utilities for reading values and formatting from worksheet cells.
 * @module cellReader
 * @category Services
 */
import AsposeCells from "aspose.cells.node";
import type { Cell, CellHorizontalAlignment, CellMerge, CellValue, CellVerticalAlignment } from "../models/Cell.ts";
import type { ColumnIndex } from "../models/Column.ts";
import type { RowIndex } from "../models/Row.ts";

/**
 * Reads the value of a cell from a worksheet.
 *
 * @param {AsposeCells.Worksheet} worksheet Worksheet instance.
 * @param {RowIndex} r Row index (0-based).
 * @param {ColumnIndex} c Column index (0-based).
 * @returns {CellValue} Value of the cell (string, number, boolean, Date, or empty string).
 */
export function readCellValue(worksheet: AsposeCells.Worksheet, r: RowIndex, c: ColumnIndex): CellValue {
	const cell = worksheet.cells.get(r, c);

	return getCellValue(cell);
}

export function readCell(worksheet: AsposeCells.Worksheet, r: RowIndex, c: ColumnIndex): Cell {
	/**
	 * Reads all properties and formatting of a cell from a worksheet.
	 *
	 * @param {AsposeCells.Worksheet} worksheet Worksheet instance.
	 * @param {RowIndex} r Row index (0-based).
	 * @param {ColumnIndex} c Column index (0-based).
	 * @returns {Cell} Cell object with value, formatting, merge, and comment info.
	 */
	const cell = worksheet.cells.get(r, c);

	const value = getCellValue(cell);
	const formula = cell.formula;

	const style = cell.getStyle();
	const font = style.font;
	const fontName = font.getName();
	const fontSize = font.size;
	const fontBold = font.isBold;
	const fontItalic = font.isItalic;
	const fontColor = decodeColor(font.color);

	const backgroundColor = decodeColor(style.foregroundColor);

	// Use getter methods if available for alignment
	const horizontalAlignment = decodeHorizontalAlignment(typeof style.getHorizontalAlignment === "function" ? style.getHorizontalAlignment() : style.horizontalAlignment);
	const verticalAlignment = decodeVerticalAlignment(typeof style.getVerticalAlignment === "function" ? style.getVerticalAlignment() : style.verticalAlignment);
	const rotationAngle = style.rotationAngle;
	const isTextWrapped = style.isTextWrapped;

	const borders = style.borders;
	const top = borders.get(AsposeCells.BorderType.TopBorder);
	const bottom = borders.get(AsposeCells.BorderType.BottomBorder);
	const left = borders.get(AsposeCells.BorderType.LeftBorder);
	const right = borders.get(AsposeCells.BorderType.RightBorder);
	const borderTopStyle = decodeCellBorderType(top.lineStyle);
	const borderTopColor = decodeColor(top.color);
	const borderBottomStyle = decodeCellBorderType(bottom.lineStyle);
	const borderBottomColor = decodeColor(bottom.color);
	const borderLeftStyle = decodeCellBorderType(left.lineStyle);
	const borderLeftColor = decodeColor(left.color);
	const borderRightStyle = decodeCellBorderType(right.lineStyle);
	const borderRightColor = decodeColor(right.color);

	const numberFormat = style.custom; // See https://reference.aspose.com/cells/nodejs-cpp/style/#custom-- and https://reference.aspose.com/cells/nodejs-cpp/style/#number--

	const isLocked = style.isLocked;

	const indentLevel = style.indentLevel;

	const shrinkToFit = style.shrinkToFit;

	const comment = cell.comment?.note || "";

	let merge: CellMerge = null;
	const mergedAreas = worksheet.cells.getMergedAreas() ?? [];

	for (const area of mergedAreas) {
		const { startRow, startColumn, endRow, endColumn } = area;
		if (!(r >= startRow && r <= endRow && c >= startColumn && c <= endColumn)) continue;
		if (r > startRow && c > startColumn) {
			merge = "up-left";
		} else if (r > startRow) {
			merge = "up";
		} else if (c > startColumn) {
			merge = "left";
		} else {
			merge = null; // Top-left cell of merged area
		}
		break;
	}

	return {
		value,
		formula,
		fontName,
		fontSize,
		fontBold,
		fontItalic,
		fontColor,
		backgroundColor,
		horizontalAlignment,
		verticalAlignment,
		rotationAngle,
		isTextWrapped,
		borderTopStyle,
		borderTopColor,
		borderBottomStyle,
		borderBottomColor,
		borderLeftStyle,
		borderLeftColor,
		borderRightStyle,
		borderRightColor,
		numberFormat,
		isLocked,
		indentLevel,
		shrinkToFit,
		merge,
		comment,
	};
}
function getCellValue(cell: AsposeCells.Cell): CellValue {
	const value = cell.value;
	if (value === null || value === undefined) return "";
	if (typeof value === "string") return value;
	if (typeof value === "number") return value;
	if (typeof value === "boolean") return value;
	if (value instanceof Date) return value;
	return "";
}
function decodeCellBorderType(lineStyle: AsposeCells.CellBorderType): Cell["borderTopStyle"] {
	switch (lineStyle) {
		case AsposeCells.CellBorderType.Thin:
			return "thin";
		case AsposeCells.CellBorderType.Medium:
			return "medium";
		case AsposeCells.CellBorderType.Thick:
			return "thick";
		case AsposeCells.CellBorderType.Dashed:
			return "dashed";
		case AsposeCells.CellBorderType.Dotted:
			return "dotted";
		case AsposeCells.CellBorderType.Double:
			return "double";
		default:
			return "thin"; // TODO
	}
}
function decodeColor(color: AsposeCells.Color): string {
	const r = color.r.toString(16).padStart(2, "0");
	const g = color.g.toString(16).padStart(2, "0");
	const b = color.b.toString(16).padStart(2, "0");
	return (r + g + b).toUpperCase();
}
function decodeHorizontalAlignment(val: AsposeCells.TextAlignmentType): CellHorizontalAlignment {
	switch (val) {
		case AsposeCells.TextAlignmentType.Left:
			return "left";
		case AsposeCells.TextAlignmentType.Center:
			return "center";
		case AsposeCells.TextAlignmentType.Right:
			return "right";
		default:
			return "left";
	}
}
function decodeVerticalAlignment(val: AsposeCells.TextAlignmentType): CellVerticalAlignment {
	switch (val) {
		case AsposeCells.TextAlignmentType.Top:
			return "top";
		case AsposeCells.TextAlignmentType.Center:
			return "middle";
		case AsposeCells.TextAlignmentType.Bottom:
			return "bottom";
		default:
			return "top";
	}
}
