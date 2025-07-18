import type { Worksheet } from "aspose.cells.node";
import { describe, expect, it } from "vitest";
import createWorkbook from "./createWorkbook";
import { deleteCells } from "./deleteCells";

const template = {
	Sheet1: [
		["A1", "B1", "C1"],
		["A2", "B2", "C2"],
		["A3", "B3", "C3"],
	],
};

describe("deleteCells", () => {
	it("can delete row", async () => {
		const workbook = await createWorkbook(template);
		const worksheet = workbook.worksheets.get(0);
		deleteCells(worksheet, "2:2");

		const values = getValues(worksheet);
		expect(values).toEqual([
			["A1", "B1", "C1"],
			["A3", "B3", "C3"],
			[null, null, null],
		]);
	});

	it("can delete column", async () => {
		const workbook = await createWorkbook(template);
		const worksheet = workbook.worksheets.get(0);
		deleteCells(worksheet, "B:B");

		const values = getValues(worksheet);
		expect(values).toEqual([
			["A1", "C1", null],
			["A2", "C2", null],
			["A3", "C3", null],
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
