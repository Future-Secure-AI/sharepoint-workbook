import type ExcelJS from "exceljs";
import type { CellValue as ExcelJSCellValue } from "exceljs";
import NotFoundError from "microsoft-graph/NotFoundError";
import type { Cell, CellValue, CellWrite } from "../models/Cell.ts";
import type { WorksheetName } from "../models/Worksheet.ts";

export function updateExcelCell(excelCell: ExcelJS.Cell, write: CellWrite): void {
	if (write.value !== undefined) excelCell.value = toExcelValue(write.value);
	if (write.format !== undefined) excelCell.numFmt = write.format ?? "";

	const alignment = mapAlignment(write);
	if (alignment) excelCell.alignment = alignment;

	const border = mapBorders(write);
	if (border) excelCell.border = border;

	const fill = mapFill(write);
	if (fill) excelCell.fill = fill;

	const font = mapFont(write);
	if (font) excelCell.font = font;
}

export function fromExcelCell(excelCell: ExcelJS.Cell): Cell {
	const value = fromExcelValue(excelCell.value);
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
		fontUnderline: excelCell.font?.underline === true ? "single" : excelCell.font?.underline === false ? "none" : (excelCell.font?.underline ?? null),
		fontStrike: excelCell.font?.strike ?? null,

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

export function toExcelValue(value: CellValue | undefined): ExcelJSCellValue {
	if (value === undefined) return "";
	if (typeof value === "string" && value.startsWith("=")) return { formula: value.slice(1) };

	return value;
}

export function fromExcelValue(excelValue: ExcelJSCellValue): CellValue {
	if (excelValue === null || excelValue === undefined) return "";

	if (typeof excelValue === "number" || typeof excelValue === "boolean" || typeof excelValue === "string" || excelValue instanceof Date) return excelValue;

	if (typeof excelValue === "object") {
		if ("formula" in excelValue && typeof excelValue.formula === "string") return `=${excelValue.formula}`;
		if ("sharedFormula" in excelValue && typeof excelValue.sharedFormula === "string") return `=${excelValue.sharedFormula}`;
		if ("hyperlink" in excelValue && typeof excelValue.hyperlink === "string") return typeof excelValue.text === "string" ? excelValue.text : excelValue.hyperlink;
		if ("richText" in excelValue && Array.isArray(excelValue.richText)) return excelValue.richText.map((rt) => rt.text).join("");
		if ("error" in excelValue) return excelValue.error;
	}

	return ""; // TODO: Think more on fallback value. Should this be an exception instead?
}

export function getWorksheetByName(workbook: ExcelJS.Workbook, worksheetName: WorksheetName): ExcelJS.Worksheet {
	const worksheet = workbook.getWorksheet(worksheetName);
	if (!worksheet) throw new NotFoundError(`Worksheet not found: ${worksheetName}`);
	return worksheet;
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
	if (cell.alignmentHorizontal && horizontalMap[cell.alignmentHorizontal]) horizontal = horizontalMap[cell.alignmentHorizontal];

	if (cell.alignmentVertical && verticalMap[cell.alignmentVertical]) vertical = verticalMap[cell.alignmentVertical];

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
function mapFill(cell: import("../models/Cell.ts").CellWrite): ExcelJS.Fill | undefined {
	if (!cell.fillForegroundColor) return undefined;
	return {
		type: "pattern",
		pattern: "solid",
		fgColor: { argb: cell.fillForegroundColor },
		bgColor: cell.fillBackgroundColor ? { argb: cell.fillBackgroundColor } : undefined,
	};
}
function mapFont(cell: import("../models/Cell.ts").CellWrite): Partial<ExcelJS.Font> | undefined {
	const result: Partial<ExcelJS.Font> = {};
	if (typeof cell.fontName === "string") result.name = cell.fontName;
	if (typeof cell.fontSize === "number") result.size = cell.fontSize;
	if (cell.fontColor) result.color = { argb: cell.fontColor };
	if (typeof cell.fontBold === "boolean") result.bold = cell.fontBold;
	if (typeof cell.fontItalic === "boolean") result.italic = cell.fontItalic;
	if (cell.fontUnderline && cell.fontUnderline !== "none") result.underline = true;
	if (typeof cell.fontStrike === "boolean") result.strike = cell.fontStrike;
	return Object.keys(result).length > 0 ? result : undefined;
}
