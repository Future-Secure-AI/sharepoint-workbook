import AsposeCells from "aspose.cells.node";
import { beforeEach, describe, expect, it } from "vitest";
import { writeCell } from "./cellWriter";

describe("writeCell", () => {
	let workbook: AsposeCells.Workbook;
	let worksheet: AsposeCells.Worksheet;

	beforeEach(() => {
		workbook = new AsposeCells.Workbook();
		worksheet = workbook.worksheets.get(0);
	});

	it("writes a string value", () => {
		writeCell(worksheet, 0, 0, "Hello");
		const cell = worksheet.cells.get(0, 0);
		expect(cell.value).toBe("Hello");
		expect(cell.getStyle().font.getName()).toBe("Arial"); // default font
	});

	it("writes a number value", () => {
		writeCell(worksheet, 1, 1, 123.45);
		const cell = worksheet.cells.get(1, 1);
		expect(cell.value).toBe(123.45);
	});

	it("writes a boolean value", () => {
		writeCell(worksheet, 2, 2, true);
		const cell = worksheet.cells.get(2, 2);
		expect(cell.value).toBe(true);
	});

	it("writes a Date value", () => {
		const date = new Date("2023-01-01T00:00:00Z");
		writeCell(worksheet, 3, 3, date);
		const cell = worksheet.cells.get(3, 3);
		expect(cell.value).toBe(44927.416666666664);
	});

	it("writes a formula string", () => {
		writeCell(worksheet, 4, 0, "=SUM(1,2)");
		const cell = worksheet.cells.get(4, 0);
		expect(cell.value).toBe(null);
		expect(cell.formula).toBe("=SUM(1,2)");
	});

	it("writes cell with formatting", () => {
		writeCell(worksheet, 5, 0, {
			value: "Styled",
			fontName: "Calibri",
			fontSize: 14,
			fontBold: true,
			fontItalic: true,
			fontColor: "FF0000",
			backgroundColor: "00FF00",
			horizontalAlignment: "right",
			verticalAlignment: "bottom",
			rotationAngle: 45,
			isTextWrapped: true,
			borderTopStyle: "thick",
			borderTopColor: "0000FF",
			borderBottomStyle: "dashed",
			borderBottomColor: "00FFFF",
			borderLeftStyle: "double",
			borderLeftColor: "FF00FF",
			borderRightStyle: "dotted",
			borderRightColor: "FFFF00",
			numberFormat: "#,##0.00",
			isLocked: false,
			indentLevel: 2,
			shrinkToFit: true,
			comment: "Test comment",
		});
		const cell = worksheet.cells.get(5, 0);
		const style = cell.getStyle();
		expect(cell.value).toBe("Styled");
		expect(style.font.getName()).toBe("Calibri");
		expect(style.font.size).toBe(14);
		expect(style.font.isBold).toBe(true);
		expect(style.font.isItalic).toBe(true);
		expect(style.font.color.r).toBe(255);
		expect(style.font.color.g).toBe(0);
		expect(style.font.color.b).toBe(0);
		expect(style.foregroundColor.r).toBe(0);
		expect(style.foregroundColor.g).toBe(255);
		expect(style.foregroundColor.b).toBe(0);
		expect(style.horizontalAlignment).toBe(AsposeCells.TextAlignmentType.Right);
		expect(style.verticalAlignment).toBe(AsposeCells.TextAlignmentType.Bottom);
		expect(style.rotationAngle).toBe(45);
		expect(style.isTextWrapped).toBe(true);
		expect(style.borders.get(AsposeCells.BorderType.TopBorder).lineStyle).toBe(AsposeCells.CellBorderType.Thick);
		expect(style.borders.get(AsposeCells.BorderType.TopBorder).color.b).toBe(255);
		expect(style.borders.get(AsposeCells.BorderType.BottomBorder).lineStyle).toBe(AsposeCells.CellBorderType.Dashed);
		expect(style.borders.get(AsposeCells.BorderType.BottomBorder).color.g).toBe(255);
		expect(style.borders.get(AsposeCells.BorderType.LeftBorder).lineStyle).toBe(AsposeCells.CellBorderType.Double);
		expect(style.borders.get(AsposeCells.BorderType.LeftBorder).color.r).toBe(255);
		expect(style.borders.get(AsposeCells.BorderType.LeftBorder).color.b).toBe(255);
		expect(style.borders.get(AsposeCells.BorderType.RightBorder).lineStyle).toBe(AsposeCells.CellBorderType.Dotted);
		expect(style.borders.get(AsposeCells.BorderType.RightBorder).color.r).toBe(255);
		expect(style.borders.get(AsposeCells.BorderType.RightBorder).color.g).toBe(255);
		expect(style.custom).toBe("#,##0.00");
		expect(style.isLocked).toBe(false);
		expect(style.indentLevel).toBe(2);
		expect(style.shrinkToFit).toBe(true);
		// Comment
		expect(cell.comment.note).toBe("Test comment");
	});

	it("merges cells up", () => {
		writeCell(worksheet, 6, 0, { value: "A", merge: "up" });
		const mergedAreas = worksheet.cells.getMergedAreas();
		expect(mergedAreas.length).toBe(1);
		expect(mergedAreas[0].startRow).toBe(5);
		expect(mergedAreas[0].endRow).toBe(6);
		expect(mergedAreas[0].startColumn).toBe(0);
		expect(mergedAreas[0].endColumn).toBe(0);
		// Check merged cell value
		expect(worksheet.cells.get(6, 0).value).toBe("A");
	});

	it("merges cells left", () => {
		writeCell(worksheet, 0, 7, { value: "B", merge: "left" });
		const mergedAreas = worksheet.cells.getMergedAreas();
		expect(mergedAreas.length).toBe(1);
		expect(mergedAreas[0].startRow).toBe(0);
		expect(mergedAreas[0].endRow).toBe(0);
		expect(mergedAreas[0].startColumn).toBe(6);
		expect(mergedAreas[0].endColumn).toBe(7);
		expect(worksheet.cells.get(0, 7).value).toBe("B");
	});

	it("merges cells up-left", () => {
		writeCell(worksheet, 8, 8, { value: "C", merge: "up-left" });
		const mergedAreas = worksheet.cells.getMergedAreas();
		expect(mergedAreas.length).toBe(1);
		expect(mergedAreas[0].startRow).toBe(7);
		expect(mergedAreas[0].endRow).toBe(8);
		expect(mergedAreas[0].startColumn).toBe(7);
		expect(mergedAreas[0].endColumn).toBe(8);
		expect(worksheet.cells.get(8, 8).value).toBe("C");
	});

	it("adds a comment to a cell", () => {
		writeCell(worksheet, 9, 9, { value: "D", comment: "A comment" });
		const cell = worksheet.cells.get(9, 9);
		expect(cell.comment.note).toBe("A comment");
	});

	it("throws on invalid color", () => {
		expect(() => writeCell(worksheet, 10, 0, { value: "bad", fontColor: "ZZZZZZ" })).toThrow();
	});

	it("throws on invalid alignment", () => {
		expect(() => writeCell(worksheet, 11, 0, { value: "bad", horizontalAlignment: "bad" as unknown as import("../models/Cell").CellHorizontalAlignment })).toThrow();
		expect(() => writeCell(worksheet, 12, 0, { value: "bad", verticalAlignment: "bad" as unknown as import("../models/Cell").CellVerticalAlignment })).toThrow();
	});
});
