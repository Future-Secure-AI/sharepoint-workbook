import { describe, expect, it } from "vitest";
import createWorkbook from "./createWorkbook";
import { readCells } from "./readCells";

describe("readCells", () => {
	it("reads the entire sheet with ':' ref", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				["A", "B", "C"],
				[1, 2, 3],
				[4, 5, 6],
			],
		});
		const ws = wb.worksheets.get(0);
		const cells = readCells(ws, ":");
		expect(extractValues(cells)).toEqual([
			["A", "B", "C"],
			[1, 2, 3],
			[4, 5, 6],
		]);
	});
	it("reads a rectangular block of cells", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				["A1", "B1", "C1"],
				["A2", "B2", "C2"],
				["A3", "B3", "C3"],
			],
		});
		const ws = wb.worksheets.get(0);
		const cells = readCells(ws, "A1:B2");
		expect(extractValues(cells)).toEqual([
			["A1", "B1"],
			["A2", "B2"],
		]);
	});

	it("reads a single cell", async () => {
		const wb = await createWorkbook({
			Sheet1: [[123, 456]],
		});
		const ws = wb.worksheets.get(0);
		const cells = readCells(ws, "B1");
		expect(extractValues(cells)).toEqual([[456]]);
	});

	it("reads full worksheet range", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				[1, 2],
				[3, 4],
			],
		});
		const ws = wb.worksheets.get(0);
		const cells = readCells(ws, "A1:B2");
		expect(extractValues(cells)).toEqual([
			[1, 2],
			[3, 4],
		]);
	});
});

function extractValues(cells: unknown[][]) {
	return cells.map((row) => row.map((cell) => (cell as { value: unknown }).value));
}
