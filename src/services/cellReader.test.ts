import AsposeCells from "aspose.cells.node";
import { beforeEach, describe, expect, it } from "vitest";
import { readCell, readCellValue } from "./cellReader";

describe("readCellValue", () => {
	let workbook: AsposeCells.Workbook;
	let worksheet: AsposeCells.Worksheet;

	beforeEach(() => {
		workbook = new AsposeCells.Workbook();
		worksheet = workbook.worksheets.get(0);
	});

	it("should return a string value", () => {
		worksheet.cells.get(0, 0).putValue("Hello");
		expect(readCellValue(worksheet, 0, 0)).toBe("Hello");
	});

	it("should return a number value", () => {
		worksheet.cells.get(1, 1).putValue(123.45);
		expect(readCellValue(worksheet, 1, 1)).toBe(123.45);
	});

	it("should return a boolean value", () => {
		worksheet.cells.get(2, 2).putValue(true);
		expect(readCellValue(worksheet, 2, 2)).toBe(true);
	});

	it("should return a Date value", () => {
		const date = new Date("2023-01-01T00:00:00Z");
		worksheet.cells.get(3, 3).putValue(date);
		const result = readCellValue(worksheet, 3, 3);
		// Accept both Date and string (Excel may store as string or number)
		if (result instanceof Date) {
			expect(result.toISOString()).toBe(date.toISOString());
		} else if (typeof result === "number") {
			// Excel serial date to JS Date
			const excelEpoch = new Date(Date.UTC(1899, 11, 30));
			const msPerDay = 24 * 60 * 60 * 1000;
			const parsed = new Date(excelEpoch.getTime() + result * msPerDay);
			expect(parsed.toISOString().slice(0, 10)).toBe(date.toISOString().slice(0, 10));
		} else if (typeof result === "string") {
			const parsed = new Date(result);
			expect(parsed.toISOString().slice(0, 10)).toBe(date.toISOString().slice(0, 10));
		} else {
			throw new Error("Unexpected result type for date cell: " + typeof result);
		}
	});

	it("should return an empty string for empty cell", () => {
		expect(readCellValue(worksheet, 4, 4)).toBe("");
	});
});

describe("readCell", () => {
	let workbook: AsposeCells.Workbook;
	let worksheet: AsposeCells.Worksheet;

	beforeEach(() => {
		workbook = new AsposeCells.Workbook();
		worksheet = workbook.worksheets.get(0);
	});

	it("should return correct value for a string cell", () => {
		worksheet.cells.get(0, 0).putValue("Test");
		const cell = readCell(worksheet, 0, 0);
		expect(cell.value).toBe("Test");
		expect(cell.formula).toBe("");
	});

	it("should return all default cell attributes for a new cell", () => {
		worksheet.cells.get(0, 0).putValue("Test");
		const cell = readCell(worksheet, 0, 0);
		expect(cell.value).toBe("Test");
		expect(cell.formula).toBe("");
		expect(cell.fontName).toBe("Arial");
		expect(cell.fontSize).toBe(10);
		expect(cell.fontBold).toBe(false);
		expect(cell.fontItalic).toBe(false);
		expect(cell.fontColor).toBe("000000");
		expect(cell.backgroundColor).toBe("000000");
		expect(cell.horizontalAlignment).toBe("left");
		expect(cell.verticalAlignment).toBe("middle");
		expect(cell.rotationAngle).toBe(0);
		expect(cell.isTextWrapped).toBe(false);
		expect(cell.borderTopStyle).toBe("thin");
		expect(cell.borderTopColor).toBe("000000");
		expect(cell.borderBottomStyle).toBe("thin");
		expect(cell.borderBottomColor).toBe("000000");
		expect(cell.borderLeftStyle).toBe("thin");
		expect(cell.borderLeftColor).toBe("000000");
		expect(cell.borderRightStyle).toBe("thin");
		expect(cell.borderRightColor).toBe("000000");
		expect(cell.numberFormat).toBe("");
		expect(cell.isLocked).toBe(true);
		expect(cell.indentLevel).toBe(0);
		expect(cell.shrinkToFit).toBe(false);
		expect(cell.merge).toBe(null);
		expect(cell.comment).toBe("");
	});

	it("should return updated cell attributes when set", () => {
		const cellObj = worksheet.cells.get(1, 1);
		cellObj.putValue("Styled");
		const style = cellObj.getStyle();
		style.font.setName("Calibri");
		style.font.size = 14;
		style.font.isBold = true;
		style.font.isItalic = true;
		style.font.color = new AsposeCells.Color(0, 0, 255); // Blue
		style.foregroundColor = new AsposeCells.Color(255, 255, 0); // Yellow
		style.horizontalAlignment = AsposeCells.TextAlignmentType.Right;
		style.verticalAlignment = AsposeCells.TextAlignmentType.Bottom;
		style.rotationAngle = 45;
		style.isTextWrapped = true;
		style.borders.get(AsposeCells.BorderType.TopBorder).lineStyle = AsposeCells.CellBorderType.Thick;
		style.borders.get(AsposeCells.BorderType.TopBorder).color = new AsposeCells.Color(0, 255, 0); // Green
		style.borders.get(AsposeCells.BorderType.BottomBorder).lineStyle = AsposeCells.CellBorderType.Dashed;
		style.borders.get(AsposeCells.BorderType.BottomBorder).color = new AsposeCells.Color(255, 0, 0); // Red
		style.borders.get(AsposeCells.BorderType.LeftBorder).lineStyle = AsposeCells.CellBorderType.Double;
		style.borders.get(AsposeCells.BorderType.LeftBorder).color = new AsposeCells.Color(0, 0, 255); // Blue
		style.borders.get(AsposeCells.BorderType.RightBorder).lineStyle = AsposeCells.CellBorderType.Dotted;
		style.borders.get(AsposeCells.BorderType.RightBorder).color = new AsposeCells.Color(255, 255, 0); // Yellow
		style.custom = "#,##0.00";
		style.isLocked = false;
		style.indentLevel = 2;
		style.shrinkToFit = true;
		cellObj.setStyle(style);

		const cell = readCell(worksheet, 1, 1);
		expect(cell.value).toBe("Styled");
		expect(cell.fontName).toBe("Calibri");
		expect(cell.fontSize).toBe(14);
		expect(cell.fontBold).toBe(true);
		expect(cell.fontItalic).toBe(true);
		expect(cell.fontColor).toBe("0000FF");
		expect(cell.backgroundColor).toBe("FFFF00");
		expect(cell.horizontalAlignment).toBe("right");
		expect(cell.verticalAlignment).toBe("bottom");
		expect(cell.rotationAngle).toBe(45);
		expect(cell.isTextWrapped).toBe(true);
		expect(cell.borderTopStyle).toBe("thick");
		expect(cell.borderTopColor).toBe("00FF00");
		expect(cell.borderBottomStyle).toBe("dashed");
		expect(cell.borderBottomColor).toBe("FF0000");
		expect(cell.borderLeftStyle).toBe("double");
		expect(cell.borderLeftColor).toBe("0000FF");
		expect(cell.borderRightStyle).toBe("dotted");
		expect(cell.borderRightColor).toBe("FFFF00");
		expect(cell.numberFormat).toBe("#,##0.00");
		expect(cell.isLocked).toBe(false);
		expect(cell.indentLevel).toBe(2);
		expect(cell.shrinkToFit).toBe(true);
	});

	it("should return empty string for empty cell", () => {
		const cell = readCell(worksheet, 4, 4);
		expect(cell.value).toBe("");
	});

	it("should return correct value for a number cell", () => {
		worksheet.cells.get(1, 0).putValue(123.45);
		const cell = readCell(worksheet, 1, 0);
		expect(cell.value).toBe(123.45);
		expect(cell.formula).toBe("");
	});

	it("should return correct value for a boolean cell", () => {
		worksheet.cells.get(2, 0).putValue(true);
		const cell = readCell(worksheet, 2, 0);
		expect(cell.value).toBe(true);
		expect(cell.formula).toBe("");
	});

	it("should return correct value for a date cell", () => {
		const date = new Date("2023-01-01T00:00:00Z");
		worksheet.cells.get(3, 0).putValue(date);
		const cell = readCell(worksheet, 3, 0);
		expect(cell.value).toBeCloseTo(44927.416666666664, 10); // TODO: Shouldn't this be a Date?
		expect(cell.formula).toBe("");
	});

	it("should detect merged cells and set merge property", () => {
		worksheet.cells.merge(5, 5, 2, 2); // Merge 5,5 to 6,6
		const topLeft = readCell(worksheet, 5, 5);
		const down = readCell(worksheet, 6, 5);
		const right = readCell(worksheet, 5, 6);
		const downRight = readCell(worksheet, 6, 6);
		expect(topLeft.merge).toBe(null);
		expect(down.merge).toBe("up");
		expect(right.merge).toBe("left");
		expect(downRight.merge).toBe("up-left");
	});

	it("should return comment if present", () => {
		worksheet.comments.add(7, 7);
		worksheet.comments.get(7, 7).note = "This is a comment";
		const cell = readCell(worksheet, 7, 7);
		expect(cell.comment).toBe("This is a comment");
	});
});
