import { describe, expect, it } from "vitest";
import createWorkbook from "./createWorkbook";
import { readCellValues } from "./readCellValues";

describe("readCellValues", () => {
	it("reads a rectangular block of cell values", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				["A1", "B1", "C1"],
				["A2", "B2", "C2"],
				["A3", "B3", "C3"],
			],
		});
		const ws = wb.worksheets.get(0);
		const values = readCellValues(ws, "A1:B2");
		expect(values).toEqual([
			["A1", "B1"],
			["A2", "B2"],
		]);
	});

	it("reads a single cell", async () => {
		const wb = await createWorkbook({
			Sheet1: [[123, 456]],
		});
		const ws = wb.worksheets.get(0);
		const values = readCellValues(ws, "B1");
		expect(values).toEqual([[456]]);
	});

	it("reads full worksheet range", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				[1, 2],
				[3, 4],
			],
		});
		const ws = wb.worksheets.get(0);
		const values = readCellValues(ws, "A1:B2");
		expect(values).toEqual([
			[1, 2],
			[3, 4],
		]);
	});

	it("reads the entire sheet with ':' ref", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				["A", "B", "C"],
				[1, 2, 3],
				[4, 5, 6],
			],
		});
		const ws = wb.worksheets.get(0);
		const values = readCellValues(ws, ":");
		expect(values).toEqual([
			["A", "B", "C"],
			[1, 2, 3],
			[4, 5, 6],
		]);
	});

	it("returns empty array for zero-height range", async () => {
		const wb = await createWorkbook({
			Sheet1: [[1, 2]],
		});
		const ws = wb.worksheets.get(0);
		expect(() => readCellValues(ws, "A2:B1")).toThrow();
	});

	it("returns empty rows for zero-width range", async () => {
		const wb = await createWorkbook({
			Sheet1: [
				[1, 2],
				[3, 4],
			],
		});
		const ws = wb.worksheets.get(0);
		expect(() => readCellValues(ws, "B1:A2")).toThrow();
	});
});
