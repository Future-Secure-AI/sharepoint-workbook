import { describe, expect, it } from "vitest";
import createWorkbook from "./createWorkbook";
import { readCellValues } from "./readCellValues";
import { updateEachCell } from "./updateEachCell";

// Tests for updateEachCell

describe("updateEachCell", () => {
	it("updates every cell in a rectangular range to a value", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				[1, 2, 3],
				[4, 5, 6],
				[7, 8, 9],
			],
		});
		const ws = wb.worksheets.get(0);
		updateEachCell(ws, "B2:C3", 0);
		const values = readCellValues(ws, "A1:C3");
		expect(values).toEqual([
			[1, 2, 3],
			[4, 0, 0],
			[7, 0, 0],
		]);
	});

	it("updates a single cell", async () => {
		const wb = await createWorkbook({
			Sheet1: [[1, 2]],
		});
		const ws = wb.worksheets.get(0);
		updateEachCell(ws, "B1", 99);
		const values = readCellValues(ws, "A1:B1");
		expect(values).toEqual([[1, 99]]);
	});

	it("updates the entire sheet with a value", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				["A", "B"],
				["C", "D"],
			],
		});
		const ws = wb.worksheets.get(0);
		updateEachCell(ws, ":", "X");
		const values = readCellValues(ws, ":");
		expect(values).toEqual([
			["X", "X"],
			["X", "X"],
		]);
	});

	it("updates with a partial cell object", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				[0, 0],
				[0, 0],
			],
		});
		const ws = wb.worksheets.get(0);
		updateEachCell(ws, "A1:B2", { value: 42 });
		const values = readCellValues(ws, "A1:B2");
		expect(values).toEqual([
			[42, 42],
			[42, 42],
		]);
	});

	it("throws for zero-height range", async () => {
		const wb = await createWorkbook({
			Sheet1: [[1, 2]],
		});
		const ws = wb.worksheets.get(0);
		expect(() => updateEachCell(ws, "A2:B1", 1)).toThrow();
	});

	it("throws for zero-width range", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				[1, 2],
				[3, 4],
			],
		});
		const ws = wb.worksheets.get(0);
		expect(() => updateEachCell(ws, "B1:A2", 1)).toThrow();
	});
});
