import { describe, expect, it } from "vitest";
import type { Workbook } from "../models/Workbook";
import createWorkbook from "./createWorkbook";
import filterWorkbookColumns from "./filterWorkbookColumns";

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

describe("filterWorkbookColumns integration", () => {
	it("filters columns", async () => {
		const handle = await createWorkbook({
			Sheet1: [
				["A", "B", "C", "D"],
				["1", "2", "3", "4"],
				["5", "6", "7", "8"],
			],
		});
		await filterWorkbookColumns(handle, (header) => header !== "B" && header !== "D");
		const rows = await readSheetRows(handle);
		expect(rows).toEqual([
			["A", "C"],
			["1", "3"],
			["5", "7"],
		]);
	});

	it("keeps all if no filter", async () => {
		const handle = await createWorkbook({
			Sheet1: [
				["A", "B"],
				["1", "2"],
			],
		});
		await filterWorkbookColumns(handle, () => true);
		const rows = await readSheetRows(handle);
		expect(rows).toEqual([
			["A", "B"],
			["1", "2"],
		]);
	});
});
