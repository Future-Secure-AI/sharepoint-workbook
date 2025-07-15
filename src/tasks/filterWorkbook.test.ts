import ExcelJS from "exceljs";
import { describe, expect, it } from "vitest";
import type { Handle } from "../models/Handle";
import { getLatestRevisionFilePath } from "../services/workingFolder";
import filterWorkbook from "./filterWorkbook";
import importWorkbook from "./importWorkbook";

describe("filterWorkbook integration", () => {
	it("filters columns and rows", async () => {
		const handle = await importWorkbook([
			{
				name: "Sheet1",
				rows: [
					["A", "B", "C", "D"],
					["1", "2", "3", "4"],
					["5", "6", "7", "8"],
				],
			},
		]);

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
		const handle = await importWorkbook([
			{
				name: "Sheet1",
				rows: [
					["skip1", "skip2"],
					["A", "B"],
					["1", "2"],
				],
			},
		]);

		await filterWorkbook(handle, { skipRows: 1 });

		const rows = await readSheetRows(handle);
		expect(rows).toEqual([
			["A", "B"],
			["1", "2"],
		]);
	});

	it("keeps all if no filters", async () => {
		const handle = await importWorkbook([
			{
				name: "Sheet1",
				rows: [
					["A", "B"],
					["1", "2"],
				],
			},
		]);

		await filterWorkbook(handle, {});

		const rows = await readSheetRows(handle);
		expect(rows).toEqual([
			["A", "B"],
			["1", "2"],
		]);
	});
});

async function readSheetRows(handle: Handle): Promise<string[][]> {
	const file = await getLatestRevisionFilePath(handle.id);
	const wb = new ExcelJS.Workbook();
	await wb.xlsx.readFile(file);
	const ws = wb.worksheets[0];
	if (!ws) return [];
	return ws
		.getSheetValues()
		.slice(1)
		.map((row) => (Array.isArray(row) ? row.slice(1).map((cell) => (cell == null ? "" : String(cell))) : []));
}
