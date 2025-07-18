import type { Worksheet } from "aspose.cells.node";
import { describe, expect, it } from "vitest";
import type { CellValue } from "../models/Cell.ts";
import { readCellValue } from "../services/cellReader.ts";
import createWorkbook from "./createWorkbook";
import insertCells from "./insertCells";

const template = {
	Sheet1: [
		["A1", "B1", "C1"],
		["A2", "B2", "C2"],
		["A3", "B3", "C3"],
	],
};

describe("insertCells", () => {
	it("can insert down", async () => {
		const workbook = await createWorkbook(template);
		const worksheet = workbook.worksheets.get(0);
		insertCells(worksheet, "B2", "down", [
			["_A1", "_B1"],
			["_A2", "_B2"],
		]);

		const values = getValues(worksheet);
		expect(values).toEqual([
			["A1", "B1", "C1", ""],
			["", "_A1", "_B1", ""],
			["", "_A2", "_B2", ""],
			["A2", "B2", "C2", ""],
		]);
	});

	it("can insert right", async () => {
		const workbook = await createWorkbook(template);
		const worksheet = workbook.worksheets.get(0);
		insertCells(worksheet, "B2", "right", [
			["_A1", "_B1"],
			["_A2", "_B2"],
		]);

		const values = getValues(worksheet);
		expect(values).toEqual([
			["A1", "", "", "B1"],
			["A2", "_A1", "_B1", "B2"],
			["A3", "_A2", "_B2", "B3"],
			["", "", "", ""],
		]);
	});
});

function getValues(worksheet: Worksheet) {
	const grid: CellValue[][] = [];
	for (let r = 0; r < 4; r++) {
		const row: CellValue[] = [];
		for (let c = 0; c < 4; c++) {
			row.push(readCellValue(worksheet, r, c));
		}
		grid.push(row);
	}
	return grid;
}
