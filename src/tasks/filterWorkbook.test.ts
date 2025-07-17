import { describe, expect, it } from "vitest";
import type { Workbook } from "../models/Workbook";
import createWorkbook from "./createWorkbook";
import filterWorkbook from "./filterWorkbook";

describe("filterWorkbook integration", () => {
	it("filters columns and rows", async () => {
		const handle = await createWorkbook({
			Sheet1: [
				["A", "B", "C", "D"],
				["1", "2", "3", "4"],
				["5", "6", "7", "8"],
			],
		});

		await filterWorkbook(handle, {
			columnFilter: (header) => header !== "B" && header !== "D",
			rowFilter: (cells) => cells[0] !== "5",
		});

		const rows = await readSheetRows(handle);
		expect(rows).toEqual([
			["A", "C"],
			["1", "3"],
		]);
	});

	it("skips header rows", async () => {
		const handle = await createWorkbook({
			Sheet1: [
				["skip1", "skip2"],
				["A", "B"],
				["1", "2"],
			],
		});

		await filterWorkbook(handle, { skipRows: 1 });

		const rows = await readSheetRows(handle);
		expect(rows).toEqual([
			["A", "B"],
			["1", "2"],
		]);
	});

	it("keeps all if no filters", async () => {
		const handle = await createWorkbook({
			Sheet1: [
				["A", "B"],
				["1", "2"],
			],
		});

		await filterWorkbook(handle, {});

		const rows = await readSheetRows(handle);
		expect(rows).toEqual([
			["A", "B"],
			["1", "2"],
		]);
	});
});

async function readSheetRows(workbook: Workbook): Promise<string[][]> {
	const ws = workbook.worksheets.get(0);
	if (!ws) return [];
	const rows: string[][] = [];
	for (let r = 0; r <= ws.cells.maxDataRow; r++) {
		const row: string[] = [];
		for (let c = 0; c <= ws.cells.maxDataColumn; c++) {
			const cell = ws.cells.get(r, c)?.value;
			row.push(cell == null ? "" : String(cell));
		}
		rows.push(row);
	}
	return rows;
}
