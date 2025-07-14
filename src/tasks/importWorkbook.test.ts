import ExcelJS from "exceljs";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import type { WriteRow } from "../models/Row";
import type { WriteWorksheet } from "../models/Worksheet";
import { getLatestRevisionFilePath } from "../services/workingFolder";
import importWorkbook from "./importWorkbook";

describe("importWorkbook integration", () => {
	let tempDir: string;
	let origEnv: string | undefined;

	beforeEach(() => {
		tempDir = fs.mkdtempSync(path.join(os.tmpdir(), "importWorkbook-test-"));
		origEnv = process.env.WORKING_FOLDER;
		process.env.WORKING_FOLDER = tempDir;
	});

	afterEach(() => {
		process.env.WORKING_FOLDER = origEnv;
		fs.rmSync(tempDir, { recursive: true, force: true });
	});

	it("creates a workbook with a single worksheet", async () => {
		const worksheets: WriteWorksheet[] = [
			{
				name: "Sheet1",
				rows: asRows([
					[1, 2, 3],
					[4, 5, 6],
				]),
			},
		];
		const hdl = await importWorkbook(worksheets);
		const file = await getLatestRevisionFilePath(hdl.id);

		const xls = new ExcelJS.Workbook();
		await xls.xlsx.readFile(file);

		const workbook = xls.getWorksheet("Sheet1");
		expect(workbook).toBeTruthy();

		if (workbook) {
			const values = workbook
				.getSheetValues()
				.slice(1)
				.map((row) => (Array.isArray(row) ? row.slice(1) : []));
			expect(values).toEqual([
				[1, 2, 3],
				[4, 5, 6],
			]);
		}
	});

	it("creates a workbook with multiple worksheets", async () => {
		const worksheets: WriteWorksheet[] = [
			{ name: "A", rows: asRows([[1], [2]]) },
			{ name: "B", rows: asRows([[3], [4]]) },
		];

		const handle = await importWorkbook(worksheets);
		const file = await getLatestRevisionFilePath(handle.id);
		expect(file).toBeTruthy();
		const wb = new ExcelJS.Workbook();
		await wb.xlsx.readFile(file);
		const wsA = wb.getWorksheet("A");
		const wsB = wb.getWorksheet("B");
		expect(wsA).toBeTruthy();
		expect(wsB).toBeTruthy();
		if (wsA) {
			expect(
				wsA
					.getSheetValues()
					.slice(1)
					.map((row) => (Array.isArray(row) ? row.slice(1) : [])),
			).toEqual([[1], [2]]);
		}
		if (wsB) {
			expect(
				wsB
					.getSheetValues()
					.slice(1)
					.map((row) => (Array.isArray(row) ? row.slice(1) : [])),
			).toEqual([[3], [4]]);
		}
	});

	it("supports async iterable worksheets and rows", async () => {
		async function* worksheetGen() {
			yield {
				name: "AsyncSheet",
				rows: asRows([
					[10, 20],
					[30, 40],
				]),
			};
		}
		const handle = await importWorkbook(worksheetGen());
		const file = await getLatestRevisionFilePath(handle.id);
		expect(file).toBeTruthy();
		const wb = new ExcelJS.Workbook();
		await wb.xlsx.readFile(file);
		const ws = wb.getWorksheet("AsyncSheet");
		expect(ws).toBeTruthy();
		if (ws) {
			expect(
				ws
					.getSheetValues()
					.slice(1)
					.map((row) => (Array.isArray(row) ? row.slice(1) : [])),
			).toEqual([
				[10, 20],
				[30, 40],
			]);
		}
	});
});

function asRows(rows: unknown[][]): AsyncIterable<WriteRow> {
	return (async function* () {
		for (const r of rows) yield r.map((cell) => ({ value: cell })) as WriteRow;
	})();
}
