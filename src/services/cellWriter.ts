import AsposeCells from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { Cell, CellHorizontalAlignment, CellValue } from "../models/Cell.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";

export function writeCell(worksheet: AsposeCells.Worksheet, r: number, c: number, cellOrValue: CellValue | DeepPartial<Cell>) {
	const cell = worksheet.cells.get(r, c);

	if (isCellValue(cellOrValue)) {
		setCellValueWithFormula(cell, cellOrValue);
		return;
	}

	const {
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
	} = cellOrValue;

	if (value !== undefined) setCellValue(cell, value);
	if (formula !== undefined) cell.formula = formula;

	const style = cell.getStyle();
	if (fontName !== undefined) style.font.setName(fontName);
	if (fontSize !== undefined) style.font.size = fontSize;
	if (fontBold !== undefined) style.font.isBold = fontBold;
	if (fontItalic !== undefined) style.font.isItalic = fontItalic;
	if (fontColor !== undefined) style.font.color = encodeColor(fontColor);

	if (backgroundColor !== undefined) style.foregroundColor = encodeColor(backgroundColor);

	if (horizontalAlignment !== undefined) style.horizontalAlignment = encodeHorizontalAlignment(horizontalAlignment);
	if (verticalAlignment !== undefined) style.verticalAlignment = encodeVerticalAlignment(verticalAlignment);
	if (rotationAngle !== undefined) style.rotationAngle = rotationAngle;
	if (isTextWrapped !== undefined) style.isTextWrapped = isTextWrapped;

	const borders = style.borders;
	if (borderTopStyle !== undefined) borders.get(AsposeCells.BorderType.TopBorder).lineStyle = encodeCellBorderType(borderTopStyle);
	if (borderTopColor !== undefined) borders.get(AsposeCells.BorderType.TopBorder).color = encodeColor(borderTopColor);
	if (borderBottomStyle !== undefined) borders.get(AsposeCells.BorderType.BottomBorder).lineStyle = encodeCellBorderType(borderBottomStyle);
	if (borderBottomColor !== undefined) borders.get(AsposeCells.BorderType.BottomBorder).color = encodeColor(borderBottomColor);
	if (borderLeftStyle !== undefined) borders.get(AsposeCells.BorderType.LeftBorder).lineStyle = encodeCellBorderType(borderLeftStyle);
	if (borderLeftColor !== undefined) borders.get(AsposeCells.BorderType.LeftBorder).color = encodeColor(borderLeftColor);
	if (borderRightStyle !== undefined) borders.get(AsposeCells.BorderType.RightBorder).lineStyle = encodeCellBorderType(borderRightStyle);
	if (borderRightColor !== undefined) borders.get(AsposeCells.BorderType.RightBorder).color = encodeColor(borderRightColor);

	if (numberFormat !== undefined) style.custom = numberFormat;
	if (isLocked !== undefined) style.isLocked = isLocked;
	if (indentLevel !== undefined) style.indentLevel = indentLevel;
	if (shrinkToFit !== undefined) style.shrinkToFit = shrinkToFit;

	cell.setStyle(style);

	if (merge === "up" || merge === "left" || merge === "up-left") {
		const mergedAreas = worksheet.cells.getMergedAreas();
		const mergeConfigs = {
			up: { dr: -1, dc: 0, rows: 2, cols: 1 },
			left: { dr: 0, dc: -1, rows: 1, cols: 2 },
			"up-left": { dr: -1, dc: -1, rows: 2, cols: 2 },
		} as const;
		const config = mergeConfigs[merge];
		const targetR = r + config.dr;
		const targetC = c + config.dc;
		let found = false;
		for (const area of mergedAreas) {
			if (targetR >= area.startRow && targetR <= area.endRow && targetC >= area.startColumn && targetC <= area.endColumn) {
				worksheet.cells.unMerge(area.startRow, area.startColumn, area.endRow - area.startRow + 1, area.endColumn - area.startColumn + 1);
				worksheet.cells.merge(Math.min(area.startRow, r), Math.min(area.startColumn, c), Math.abs(area.endRow - area.startRow + 1 + (r < area.startRow ? 1 : 0)), Math.abs(area.endColumn - area.startColumn + 1 + (c < area.startColumn ? 1 : 0)));
				found = true;
				break;
			}
		}
		if (!found) {
			worksheet.cells.merge(targetR, targetC, config.rows, config.cols);
		}
	}

	if (comment !== undefined) {
		if (!cell.comment) {
			worksheet.comments.add(cell.row, cell.column);
			const comm = worksheet.comments.get(cell.row, cell.column);
			comm.note = comment;
		} else {
			cell.comment.note = comment;
		}
	}
}
function isCellValue(cellOrValue: string | number | boolean | Date | DeepPartial<Cell>) {
	return typeof cellOrValue === "string" || typeof cellOrValue === "number" || typeof cellOrValue === "boolean" || cellOrValue instanceof Date;
}

function encodeCellBorderType(style: string): AsposeCells.CellBorderType {
	switch (style) {
		case "thin":
			return AsposeCells.CellBorderType.Thin;
		case "medium":
			return AsposeCells.CellBorderType.Medium;
		case "thick":
			return AsposeCells.CellBorderType.Thick;
		case "dashed":
			return AsposeCells.CellBorderType.Dashed;
		case "dotted":
			return AsposeCells.CellBorderType.Dotted;
		case "double":
			return AsposeCells.CellBorderType.Double;
		default:
			return AsposeCells.CellBorderType.Thin;
	}
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

function setCellValueWithFormula(cell: AsposeCells.Cell, value: CellValue) {
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

function encodeHorizontalAlignment(val: CellHorizontalAlignment): AsposeCells.TextAlignmentType {
	switch (val) {
		case "left":
			return AsposeCells.TextAlignmentType.Left;
		case "center":
			return AsposeCells.TextAlignmentType.Center;
		case "right":
			return AsposeCells.TextAlignmentType.Right;
		default:
			throw new InvalidArgumentError(`Invalid horizontal alignment: '${val}'. Expected 'left', 'center', or 'right'.`);
	}
}
function encodeVerticalAlignment(val: Cell["verticalAlignment"]): AsposeCells.TextAlignmentType {
	switch (val) {
		case "top":
			return AsposeCells.TextAlignmentType.Top;
		case "middle":
			return AsposeCells.TextAlignmentType.Center;
		case "bottom":
			return AsposeCells.TextAlignmentType.Bottom;
		default:
			throw new InvalidArgumentError(`Invalid vertical alignment: '${val}'. Expected 'top', 'middle', or 'bottom'.`);
	}
}
