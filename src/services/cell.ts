import AsposeCells from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { Cell, CellMerge, CellValue } from "../models/Cell.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";

export function readCellValue(worksheet: AsposeCells.Worksheet, r: number, c: number): CellValue {
	const cell = worksheet.cells.get(r, c);

	return getCellValue(cell);
}

export function readCell(worksheet: AsposeCells.Worksheet, r: number, c: number): Cell {
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

	const horizontalAlignment = decodeHorizontalAlignment(style.horizontalAlignment);
	const verticalAlignment = decodeVerticalAlignment(style.verticalAlignment);
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
	const mergedAreas = worksheet.cells.getMergedAreas();
	for (const { startRow, startColumn, endRow, endColumn } of mergedAreas) {
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

export function writeCell(worksheet: AsposeCells.Worksheet, r: number, c: number, cellOrValue: CellValue | DeepPartial<Cell>) {
	const output = worksheet.cells.get(r, c);

	// TODO. Make sure to use `setCellValue` if `cellOrValue` is a `DeepPartial<Cell>` or `setAmbigiousCellValue` if `cellOrValue` is a `CellValue`
}

function setCellValue(cell: AsposeCells.Cell, value: CellValue) {
	if (typeof value === "number") {
		cell.putValue(value);
	} else if (typeof value === "boolean") {
		cell.putValue(value);
	} else if (typeof value === "string") {
		cell.putValue(value);
	} else if (value instanceof Date) {
		cell.putValue(value); // Aspose will handle Date properly
	} else {
		throw new InvalidArgumentError(`Unsupported cell value type: ${typeof value}. Expected number, boolean, string, or Date.`);
	}
}

function setAmbigiousCellValue(cell: AsposeCells.Cell, value: CellValue) {
	if (typeof value === "number") {
		cell.putValue(value);
	} else if (typeof value === "boolean") {
		cell.putValue(value);
	} else if (typeof value === "string") {
		if (value.startsWith("=")) {
			cell.putValue("");
			cell.formula = value;
		} else {
			cell.putValue(value);
		}
	} else if (value instanceof Date) {
		cell.putValue(value); // Aspose will handle Date properly
	} else {
		throw new InvalidArgumentError(`Unsupported cell value type: ${typeof value}. Expected number, boolean, string, or Date.`);
	}
}
function getCellValue(cell: AsposeCells.Cell): CellValue {
	const value = cell.value;
	if (value.isString()) return value.toString();
	if (value.isNumber()) return value.toNumber();
	if (value.isBool()) return value.toBool();
	if (value.isDate()) return value.toDate();

	return "";
}

function encodeColor(hex: string): AsposeCells.Color {
	let color = hex.trim();
	if (color.startsWith("#")) {
		color = color.slice(1);
	}

	if (!/^[0-9a-fA-F]{6}$/.test(color)) {
		throw new InvalidArgumentError(`Invalid color string: '${hex}'. Expected a 6-digit hexadecimal string optionally prefixed with '#'.`);
	}

	const r = parseInt(color.slice(0, 2), 16);
	const g = parseInt(color.slice(2, 4), 16);
	const b = parseInt(color.slice(4, 6), 16);
	return new AsposeCells.Color(r, g, b);
}

function decodeColor(color: AsposeCells.Color): string {
	const r = color.r.toString(16).padStart(2, "0");
	const g = color.g.toString(16).padStart(2, "0");
	const b = color.b.toString(16).padStart(2, "0");
	return (r + g + b).toUpperCase();
}

function decodeHorizontalAlignment(val: AsposeCells.TextAlignmentType): Cell["horizontalAlignment"] {
	switch (val) {
		case AsposeCells.TextAlignmentType.Left:
			return "left";
		case AsposeCells.TextAlignmentType.Center:
			return "center";
		case AsposeCells.TextAlignmentType.Right:
			return "right";
		default:
			return "left"; // TODO
	}
}
function decodeVerticalAlignment(val: AsposeCells.TextAlignmentType): Cell["verticalAlignment"] {
	switch (val) {
		case AsposeCells.TextAlignmentType.Top:
			return "top";
		case AsposeCells.TextAlignmentType.Center:
			return "middle";
		case AsposeCells.TextAlignmentType.Bottom:
			return "bottom";
		default:
			return "top"; // TODO
	}
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
