import { describe, expect, it } from "vitest";
import type { Workbook } from "../models/Workbook";
import createWorkbook from "./createWorkbook";
import filterWorkbookRows from "./filterWorkbookRows";

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

describe("filterWorkbookRows integration", () => {
	it("filters rows", async () => {
		const handle = await createWorkbook({
			Sheet1: [
				["A", "B", "C"],
				["1", "2", "3"],
				["5", "6", "7"],
			],
		});
		await filterWorkbookRows(handle, (cells) => cells[0] !== "5");
		const rows = await readSheetRows(handle);
		expect(rows).toEqual([
			["A", "B", "C"],
			["1", "2", "3"],
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
		await filterWorkbookRows(handle, () => true, { skipRows: 1 });
		const rows = await readSheetRows(handle);
		expect(rows).toEqual([
			["A", "B"],
			["1", "2"],
		]);
	});

	it("keeps all if no filter", async () => {
		const handle = await createWorkbook({
			Sheet1: [
				["A", "B"],
				["1", "2"],
			],
		});
		await filterWorkbookRows(handle, () => true);
		const rows = await readSheetRows(handle);
		expect(rows).toEqual([
			["A", "B"],
			["1", "2"],
		]);
	});
});
