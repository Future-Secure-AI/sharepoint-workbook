import { describe, expect, it } from "vitest";
import createWorkbook from "./createWorkbook";
import { readCellValues } from "./readCellValues";
import { updateCells } from "./updateCells";

describe("updateCells", () => {
	it("updates a rectangular block of cell values", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				["A1", "B1", "C1"],
				["A2", "B2", "C2"],
				["A3", "B3", "C3"],
			],
		});
		const ws = wb.worksheets.get(0);
		updateCells(ws, "B2", [
			["X", "Y"],
			["Z", "W"],
		]);
		const values = readCellValues(ws, "A1:C3");
		expect(values).toEqual([
			["A1", "B1", "C1"],
			["A2", "X", "Y"],
			["A3", "Z", "W"],
		]);
	});

	it("updates a single cell", async () => {
		const wb = await createWorkbook({
			Sheet1: [[1, 2]],
		});
		const ws = wb.worksheets.get(0);
		updateCells(ws, "B1", [[99]]);
		const values = readCellValues(ws, "A1:B1");
		expect(values).toEqual([[1, 99]]);
	});

	it("updates with partial cell objects", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				[0, 0],
				[0, 0],
			],
		});
		const ws = wb.worksheets.get(0);
		updateCells(ws, "A1", [
			[{ value: 10 }, { value: 20 }],
			[{ value: 30 }, { value: 40 }],
		]);
		const values = readCellValues(ws, "A1:B2");
		expect(values).toEqual([
			[10, 20],
			[30, 40],
		]);
	});

	it("skips undefined cells", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				[1, 2],
				[3, 4],
			],
		});
		const ws = wb.worksheets.get(0);
		updateCells(ws, "A1", [[undefined, 99], [88]]);
		const values = readCellValues(ws, "A1:B2");
		expect(values).toEqual([
			[1, 99],
			[88, 4],
		]);
	});
});
