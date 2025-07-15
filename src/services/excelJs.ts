import type ExcelJS from "exceljs";
import InvalidArgumentError from "microsoft-graph/dist/cjs/errors/InvalidArgumentError";
import type { CellValue, CellWrite } from "../models/Cell.ts";
import type { RowWrite } from "../models/Row.ts";

export function appendRow(worksheet: ExcelJS.Worksheet, row: RowWrite): void {
	const cells = row.map(normalizeCell);
	const outRow = worksheet.addRow(cells.map((c) => c.value));

	cells.forEach((cell, i) => {
		const outCell = outRow.getCell(i + 1);
		if (cell.format !== undefined) outCell.numFmt = cell.format ?? "";
		if ((cell.mergeDown ?? 0) > 0 || (cell.mergeRight ?? 0) > 0) {
			const startCol = i + 1;
			const startRow = outRow.number;
			worksheet.mergeCells(startRow, startCol, startRow + (cell.mergeDown ?? 0), startCol + (cell.mergeRight ?? 0));
		}
		const alignment = mapAlignment(cell);
		if (alignment) outCell.alignment = alignment;

		const border = mapBorders(cell);
		if (border) outCell.border = border;

		const fill = mapFill(cell);
		if (fill) outCell.fill = fill;

		const font = mapFont(cell);
		if (font) outCell.font = font;
	});
	outRow.commit();
}
/**
 * Returns the ExcelJS.Cell at the given row and column (1-based indices).
 * @param worksheet The ExcelJS.Worksheet instance
 * @param row The row number (1-based)
 * @param column The column number (1-based)
 * @returns The ExcelJS.Cell object
 */
import type { Cell } from "../models/Cell.ts";
import type { ColumnNumber, RowNumber } from "../models/Reference.ts";

/**
 * Returns a Cell (our model) at the given row and column (1-based indices).
 * @param worksheet The ExcelJS.Worksheet instance
 * @param row The row number (1-based)
 * @param column The column number (1-based)
 * @returns The Cell object (from models/Cell.ts)
 */
export function getCell(worksheet: ExcelJS.Worksheet, row: RowNumber, column: ColumnNumber): Cell {
	const excelCell = worksheet.getRow(row).getCell(column);
	// Map ExcelJS.Cell to our Cell type
	// Only allow value if it matches CellValue (string | number | boolean | Date)
	let value: Cell["value"] = null;
	if (typeof excelCell.value === "string" || typeof excelCell.value === "number" || typeof excelCell.value === "boolean" || excelCell.value instanceof Date) {
		value = excelCell.value;
	}
	// Note: If note is a string, use it; if it's a Comment, use its text property if available
	let note: string | null = null;
	if (typeof excelCell.note === "string") {
		note = excelCell.note;
	} else if (excelCell.note && typeof excelCell.note === "object" && "text" in excelCell.note) {
		note = (excelCell.note as { text?: string }).text ?? null;
	}
	return {
		value,
		text: typeof excelCell.text === "string" ? excelCell.text : String(excelCell.text ?? ""),
		format: excelCell.numFmt ?? null,
		note,

		fontName: excelCell.font?.name ?? null,
		fontSize: excelCell.font?.size ?? null,
		fontFamily: excelCell.font?.family ?? null,
		fontColor: excelCell.font?.color?.argb ?? null,
		fontBold: excelCell.font?.bold ?? null,
		fontItalic: excelCell.font?.italic ?? null,
		fontUnderline: excelCell.font?.underline ? "single" : "none",
		fontStrike: excelCell.font?.strike ?? null,
		fontOutline: null, // ExcelJS does not support outline

		alignmentHorizontal: excelCell.alignment?.horizontal ?? null,
		alignmentVertical: excelCell.alignment?.vertical ?? null,
		alignmentWrapText: excelCell.alignment?.wrapText ?? null,
		alignmentShrinkToFit: excelCell.alignment?.shrinkToFit ?? null,
		alignmentIndent: excelCell.alignment?.indent ?? null,
		alignmentTextRotation: excelCell.alignment?.textRotation ?? null,

		borderTopStyle: excelCell.border?.top?.style ?? null,
		borderTopColor: excelCell.border?.top?.color?.argb ?? null,
		borderLeftStyle: excelCell.border?.left?.style ?? null,
		borderLeftColor: excelCell.border?.left?.color?.argb ?? null,
		borderBottomStyle: excelCell.border?.bottom?.style ?? null,
		borderBottomColor: excelCell.border?.bottom?.color?.argb ?? null,
		borderRightStyle: excelCell.border?.right?.style ?? null,
		borderRightColor: excelCell.border?.right?.color?.argb ?? null,

		fillForegroundColor: excelCell.fill && "fgColor" in excelCell.fill && excelCell.fill.fgColor ? (excelCell.fill.fgColor.argb ?? null) : null,
		fillBackgroundColor: excelCell.fill && "bgColor" in excelCell.fill && excelCell.fill.bgColor ? (excelCell.fill.bgColor.argb ?? null) : null,

		protectionLocked: excelCell.protection?.locked ?? null,
		protectionHidden: excelCell.protection?.hidden ?? null,
	};
}

function normalizeCell(cell: CellValue | CellWrite): CellWrite {
	const type = typeof cell;
	if (type === "string" || type === "number" || type === "boolean" || cell instanceof Date || cell === null) {
		return {
			value: cell as string | number | boolean | Date | null,
		};
	}
	if (type !== "object") {
		throw new InvalidArgumentError(`Unsupported cell type '${type}'.`);
	}

	return cell as CellWrite;
}

function mapAlignment(cell: import("../models/Cell.ts").CellWrite): Partial<ExcelJS.Alignment> | undefined {
	const horizontalMap: Record<string, ExcelJS.Alignment["horizontal"]> = {
		left: "left",
		center: "center",
		right: "right",
		fill: "fill",
		justify: "justify",
		centerContinuous: "centerContinuous",
		distributed: "distributed",
	};
	const verticalMap: Record<string, ExcelJS.Alignment["vertical"]> = {
		top: "top",
		middle: "middle",
		bottom: "bottom",
		justify: "justify",
		distributed: "distributed",
	};
	let horizontal: ExcelJS.Alignment["horizontal"] | undefined;
	let vertical: ExcelJS.Alignment["vertical"] | undefined;
	if (cell.alignmentHorizontal && horizontalMap[cell.alignmentHorizontal]) {
		horizontal = horizontalMap[cell.alignmentHorizontal];
	}
	if (cell.alignmentVertical && verticalMap[cell.alignmentVertical]) {
		vertical = verticalMap[cell.alignmentVertical];
	}
	const result: Partial<ExcelJS.Alignment> = {};
	if (horizontal) result.horizontal = horizontal;
	if (vertical) result.vertical = vertical;
	if (typeof cell.alignmentWrapText === "boolean") result.wrapText = cell.alignmentWrapText;
	if (typeof cell.alignmentIndent === "number") result.indent = cell.alignmentIndent;
	if (typeof cell.alignmentTextRotation === "number" || cell.alignmentTextRotation === "vertical") result.textRotation = cell.alignmentTextRotation;
	return Object.keys(result).length > 0 ? result : undefined;
}
function mapBorders(cell: import("../models/Cell.ts").CellWrite): Partial<ExcelJS.Borders> | undefined {
	const result: Partial<ExcelJS.Borders> = {};
	if (cell.borderTopStyle || cell.borderTopColor) {
		result.top = {};
		if (cell.borderTopStyle) result.top.style = cell.borderTopStyle as ExcelJS.BorderStyle;
		if (cell.borderTopColor) result.top.color = { argb: cell.borderTopColor };
	}
	if (cell.borderBottomStyle || cell.borderBottomColor) {
		result.bottom = {};
		if (cell.borderBottomStyle) result.bottom.style = cell.borderBottomStyle as ExcelJS.BorderStyle;
		if (cell.borderBottomColor) result.bottom.color = { argb: cell.borderBottomColor };
	}
	if (cell.borderLeftStyle || cell.borderLeftColor) {
		result.left = {};
		if (cell.borderLeftStyle) result.left.style = cell.borderLeftStyle as ExcelJS.BorderStyle;
		if (cell.borderLeftColor) result.left.color = { argb: cell.borderLeftColor };
	}
	if (cell.borderRightStyle || cell.borderRightColor) {
		result.right = {};
		if (cell.borderRightStyle) result.right.style = cell.borderRightStyle as ExcelJS.BorderStyle;
		if (cell.borderRightColor) result.right.color = { argb: cell.borderRightColor };
	}
	return Object.keys(result).length > 0 ? result : undefined;
}
function colorToARGB(color: string): string {
	// Only allow string color for now
	if (typeof color === "string") return color;
	throw new InvalidArgumentError("Unsupported color type for ExcelJS");
}
function mapFill(cell: import("../models/Cell.ts").CellWrite): ExcelJS.Fill | undefined {
	if (!cell.fillForegroundColor) return undefined;
	return {
		type: "pattern",
		pattern: "solid",
		fgColor: { argb: colorToARGB(cell.fillForegroundColor) },
		bgColor: cell.fillBackgroundColor ? { argb: colorToARGB(cell.fillBackgroundColor) } : undefined,
	};
}
function mapFont(cell: import("../models/Cell.ts").CellWrite): Partial<ExcelJS.Font> | undefined {
	const result: Partial<ExcelJS.Font> = {};
	if (typeof cell.fontName === "string") result.name = cell.fontName;
	if (typeof cell.fontSize === "number") result.size = cell.fontSize;
	if (cell.fontColor) result.color = { argb: colorToARGB(cell.fontColor) };
	if (typeof cell.fontBold === "boolean") result.bold = cell.fontBold;
	if (typeof cell.fontItalic === "boolean") result.italic = cell.fontItalic;
	if (cell.fontUnderline && cell.fontUnderline !== "none") result.underline = true;
	if (typeof cell.fontStrike === "boolean") result.strike = cell.fontStrike;
	return Object.keys(result).length > 0 ? result : undefined;
}
