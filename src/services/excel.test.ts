import type ExcelJS from "exceljs";
import { beforeEach, describe, expect, it } from "vitest";
import { fromExcelCell, fromExcelValue, toExcelValue, updateExcelCell } from "./excel";

describe("toExcelValue", () => {
	it("returns empty string for undefined", () => {
		expect(toExcelValue(undefined)).toBe("");
	});
	it("returns formula object for string starting with =", () => {
		expect(toExcelValue("=SUM(A1:A2)")).toEqual({ formula: "SUM(A1:A2)" });
	});
	it("returns value as-is for other types", () => {
		expect(toExcelValue(42)).toBe(42);
		expect(toExcelValue(true)).toBe(true);
		const date = new Date();
		expect(toExcelValue(date)).toBe(date);
		expect(toExcelValue("hello")).toBe("hello");
	});
});

describe("fromExcelValue", () => {
	it("returns primitive values as-is", () => {
		expect(fromExcelValue(123)).toBe(123);
		expect(fromExcelValue(false)).toBe(false);
		expect(fromExcelValue("abc")).toBe("abc");
		const date = new Date();
		expect(fromExcelValue(date)).toBe(date);
	});
	it("returns formula string for formula object", () => {
		expect(fromExcelValue({ formula: "A1+B1" })).toBe("=A1+B1");
	});
	it("returns sharedFormula string for sharedFormula object", () => {
		expect(fromExcelValue({ sharedFormula: "A1+B1" })).toBe("=A1+B1");
	});
	it("returns text for hyperlink object", () => {
		expect(fromExcelValue({ hyperlink: "http://x", text: "Click" })).toBe("Click");
	});
	it("returns joined text for richText array", () => {
		expect(fromExcelValue({ richText: [{ text: "A" }, { text: "B" }] })).toBe("AB");
	});
	it("returns error for error object", () => {
		expect(fromExcelValue({ error: "#DIV/0!" })).toBe("#DIV/0!");
	});
});

describe("updateExcelCell", () => {
	let cell: ExcelJS.Cell;
	beforeEach(() => {
		cell = createMockCell();
	});
	it("sets value and format", () => {
		updateExcelCell(cell, { value: 5, format: "0.00" });
		expect(cell.value).toBe(5);
		expect(cell.numFmt).toBe("0.00");
	});
	it("sets alignment, border, fill, font", () => {
		updateExcelCell(cell, {
			alignmentHorizontal: "center",
			alignmentVertical: "middle",
			borderTopStyle: "thin",
			borderTopColor: "FF0000FF",
			fillForegroundColor: "FF00FF00",
			fillBackgroundColor: "FFFF0000",
			fontName: "Arial",
			fontSize: 12,
			fontColor: "FF000000",
			fontBold: true,
			fontItalic: true,
			fontUnderline: "single",
			fontStrike: true,
		});
		expect(cell.alignment).toMatchObject({ horizontal: "center", vertical: "middle" });
		expect(cell.border?.top?.style).toBe("thin");
		expect(cell.border?.top?.color?.argb).toBe("FF0000FF");
		expect(cell.fill).toMatchObject({ fgColor: { argb: "FF00FF00" }, bgColor: { argb: "FFFF0000" } });
		expect(cell.font).toMatchObject({ name: "Arial", size: 12, color: { argb: "FF000000" }, bold: true, italic: true, underline: true, strike: true });
	});
});

describe("fromExcelCell", () => {
	it("extracts all properties from ExcelJS.Cell", () => {
		const excelCell = createMockCell({
			value: 42,
			text: "42",
			numFmt: "0.00",
			note: "note",
			font: {
				name: "Arial",
				size: 10,
				family: 2,
				color: { argb: "FF000000" },
				bold: true,
				italic: false,
				underline: true,
				strike: false,
			} as ExcelJS.Font,
			alignment: {
				horizontal: "left",
				vertical: "top",
				wrapText: true,
				shrinkToFit: false,
				indent: 1,
				textRotation: 0,
			} as ExcelJS.Alignment,
			border: {
				top: { style: "thin", color: { argb: "FF0000FF" } },
				left: { style: "dotted", color: { argb: "FF00FF00" } },
				bottom: { style: "double", color: { argb: "FFFF0000" } },
				right: { style: "thick", color: { argb: "FF00FFFF" } },
			} as ExcelJS.Borders,
			fill: {
				fgColor: { argb: "FF00FF00" },
				bgColor: { argb: "FFFF0000" },
				type: "pattern",
				pattern: "solid",
			} as ExcelJS.Fill,
			protection: { locked: true, hidden: false } as ExcelJS.Protection,
		});
		const cell = fromExcelCell(excelCell);
		expect(cell.value).toBe(42);
		expect(cell.text).toBe("42");
		expect(cell.format).toBe("0.00");
		expect(cell.note).toBe("note");
		expect(cell.fontName).toBe("Arial");
		expect(cell.fontSize).toBe(10);
		expect(cell.fontFamily).toBe(2);
		expect(cell.fontColor).toBe("FF000000");
		expect(cell.fontBold).toBe(true);
		expect(cell.fontItalic).toBe(false);
		expect(cell.fontUnderline).toBe("single");
		expect(cell.fontStrike).toBe(false);
		expect(cell.alignmentHorizontal).toBe("left");
		expect(cell.alignmentVertical).toBe("top");
		expect(cell.alignmentWrapText).toBe(true);
		expect(cell.alignmentShrinkToFit).toBe(false);
		expect(cell.alignmentIndent).toBe(1);
		expect(cell.alignmentTextRotation).toBe(0);
		expect(cell.borderTopStyle).toBe("thin");
		expect(cell.borderTopColor).toBe("FF0000FF");
		expect(cell.borderLeftStyle).toBe("dotted");
		expect(cell.borderLeftColor).toBe("FF00FF00");
		expect(cell.borderBottomStyle).toBe("double");
		expect(cell.borderBottomColor).toBe("FFFF0000");
		expect(cell.borderRightStyle).toBe("thick");
		expect(cell.borderRightColor).toBe("FF00FFFF");
		expect(cell.fillForegroundColor).toBe("FF00FF00");
		expect(cell.fillBackgroundColor).toBe("FFFF0000");
		expect(cell.protectionLocked).toBe(true);
		expect(cell.protectionHidden).toBe(false);
	});
	it("handles missing/undefined properties", () => {
		const excelCell = createMockCell({ value: 1 });
		const cell = fromExcelCell(excelCell);
		expect(cell.value).toBe(1);
		expect(cell.format).toBeNull();
		expect(cell.note).toBeNull();
	});
});

function createMockCell(overrides: Partial<ExcelJS.Cell> = {}): ExcelJS.Cell {
	return {
		value: undefined,
		numFmt: undefined,
		note: undefined,
		text: undefined,
		font: undefined,
		alignment: undefined,
		border: undefined,
		fill: undefined,
		protection: undefined,
		...overrides,
	} as unknown as ExcelJS.Cell;
}
