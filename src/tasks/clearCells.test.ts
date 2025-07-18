import type { Worksheet } from "aspose.cells.node";
import { describe, expect, it } from "vitest";
import { clearCells } from "./clearCells";
import createWorkbook from "./createWorkbook";

const template = {
	Sheet1: [
		["A1", "B1", "C1"],
		["A2", "B2", "C2"],
		["A3", "B3", "C3"],
	],
};

describe("clearCells", () => {
	it("can clear range", async () => {
		const workbook = await createWorkbook(template);
		const worksheet = workbook.worksheets.get(0);
		clearCells(worksheet, "A1:B2");

		const values = getValues(worksheet);
		expect(values).toEqual([
			[null, null, "C1"],
			[null, null, "C2"],
			["A3", "B3", "C3"],
		]);
	});

	it("can clear cell", async () => {
		const workbook = await createWorkbook(template);
		const worksheet = workbook.worksheets.get(0);
		clearCells(worksheet, "B2");

		const values = getValues(worksheet);
		expect(values).toEqual([
			["A1", "B1", "C1"],
			["A2", null, "C2"],
			["A3", "B3", "C3"],
		]);
	});

	it("can clear row", async () => {
		const workbook = await createWorkbook(template);
		const worksheet = workbook.worksheets.get(0);
		clearCells(worksheet, "2");

		const values = getValues(worksheet);
		expect(values).toEqual([
			["A1", "B1", "C1"],
			[null, null, null],
			["A3", "B3", "C3"],
		]);
	});

	it("can clear column", async () => {
		const workbook = await createWorkbook(template);
		const worksheet = workbook.worksheets.get(0);
		clearCells(worksheet, "B");

		const values = getValues(worksheet);
		expect(values).toEqual([
			["A1", null, "C1"],
			["A2", null, "C2"],
			["A3", null, "C3"],
		]);
	});
});

function getValues(worksheet: Worksheet) {
	return [
		[worksheet.cells.get(0, 0).value, worksheet.cells.get(0, 1).value, worksheet.cells.get(0, 2).value],
		[worksheet.cells.get(1, 0).value, worksheet.cells.get(1, 1).value, worksheet.cells.get(1, 2).value],
		[worksheet.cells.get(2, 0).value, worksheet.cells.get(2, 1).value, worksheet.cells.get(2, 2).value],
	];
}
