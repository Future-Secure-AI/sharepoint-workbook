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

	it("should return correct value and formatting for a string cell", () => {
		worksheet.cells.get(0, 0).putValue("Test");
		const cell = readCell(worksheet, 0, 0);
		expect(cell.value).toBe("Test");
		expect(cell.formula).toBe("");
		expect(typeof cell.fontName).toBe("string");
		expect(typeof cell.fontSize).toBe("number");
		expect(typeof cell.fontBold).toBe("boolean");
		expect(typeof cell.fontItalic).toBe("boolean");
		expect(typeof cell.fontColor).toBe("string");
		expect(typeof cell.backgroundColor).toBe("string");
		expect(["left", "center", "right"]).toContain(cell.horizontalAlignment);
		expect(["top", "middle", "bottom"]).toContain(cell.verticalAlignment);
		expect(typeof cell.rotationAngle).toBe("number");
		expect(typeof cell.isTextWrapped).toBe("boolean");
		expect(["thin", "medium", "thick", "dashed", "dotted", "double"]).toContain(cell.borderTopStyle);
		expect(typeof cell.borderTopColor).toBe("string");
		expect(["thin", "medium", "thick", "dashed", "dotted", "double"]).toContain(cell.borderBottomStyle);
		expect(typeof cell.borderBottomColor).toBe("string");
		expect(["thin", "medium", "thick", "dashed", "dotted", "double"]).toContain(cell.borderLeftStyle);
		expect(typeof cell.borderLeftColor).toBe("string");
		expect(["thin", "medium", "thick", "dashed", "dotted", "double"]).toContain(cell.borderRightStyle);
		expect(typeof cell.borderRightColor).toBe("string");
		expect(typeof cell.numberFormat).toBe("string");
		expect(typeof cell.isLocked).toBe("boolean");
		expect(typeof cell.indentLevel).toBe("number");
		expect(typeof cell.shrinkToFit).toBe("boolean");
		expect(["up", "left", "up-left", null]).toContain(cell.merge);
		expect(typeof cell.comment).toBe("string");
	});

	it("should return correct value for a number cell", () => {
		worksheet.cells.get(1, 1).putValue(42);
		const cell = readCell(worksheet, 1, 1);
		expect(cell.value).toBe(42);
	});

	it("should return correct value for a boolean cell", () => {
		worksheet.cells.get(2, 2).putValue(true);
		const cell = readCell(worksheet, 2, 2);
		expect(cell.value).toBe(true);
	});

	it("should return correct value for a date cell", () => {
		const date = new Date("2023-01-01T00:00:00Z");
		worksheet.cells.get(3, 3).putValue(date);
		const cell = readCell(worksheet, 3, 3);
		if (cell.value instanceof Date) {
			expect(cell.value.toISOString()).toBe(date.toISOString());
		} else if (typeof cell.value === "number") {
			const excelEpoch = new Date(Date.UTC(1899, 11, 30));
			const msPerDay = 24 * 60 * 60 * 1000;
			const parsed = new Date(excelEpoch.getTime() + cell.value * msPerDay);
			expect(parsed.toISOString().slice(0, 10)).toBe(date.toISOString().slice(0, 10));
		} else if (typeof cell.value === "string") {
			const parsed = new Date(cell.value);
			expect(parsed.toISOString().slice(0, 10)).toBe(date.toISOString().slice(0, 10));
		} else {
			throw new Error("Unexpected result type for date cell: " + typeof cell.value);
		}
	});

	it("should return empty string for empty cell", () => {
		const cell = readCell(worksheet, 4, 4);
		expect(cell.value).toBe("");
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
