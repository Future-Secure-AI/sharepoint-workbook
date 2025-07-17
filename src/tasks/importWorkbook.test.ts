import { describe, expect, it } from "vitest";
import importWorkbook from "./importWorkbook";

describe("importWorkbook", () => {
	it("can import single worksheet", async () => {
		const handle = await importWorkbook({
			Sheet1: [
				[1, 2, 3],
				[4, 5, 6],
			],
		});

		const workbook = handle.workbook;
		const worksheet = workbook.worksheets.get("Sheet1");
		expect(worksheet).toBeTruthy();
		if (worksheet) {
			const values: unknown[][] = [];
			for (let r = 0; r < worksheet.cells.maxDataRow + 1; r++) {
				const row: unknown[] = [];
				for (let c = 0; c < worksheet.cells.maxDataColumn + 1; c++) {
					row.push(worksheet.cells.get(r, c).value);
				}
				values.push(row);
			}
			expect(values).toEqual([
				[1, 2, 3],
				[4, 5, 6],
			]);
		}
	});

	it("can import multiple worksheets", async () => {
		const handle = await importWorkbook({
			A: [[1], [2]],
			B: [[3], [4]],
		});

		const workbook = handle.workbook;
		const wsA = workbook.worksheets.get("A");
		const wsB = workbook.worksheets.get("B");
		expect(wsA).toBeTruthy();
		expect(wsB).toBeTruthy();
		if (wsA) {
			const valuesA: unknown[][] = [];
			for (let r = 0; r < wsA.cells.maxDataRow + 1; r++) {
				const row: unknown[] = [];
				for (let c = 0; c < wsA.cells.maxDataColumn + 1; c++) {
					row.push(wsA.cells.get(r, c).value);
				}
				valuesA.push(row);
			}
			expect(valuesA).toEqual([[1], [2]]);
		}
		if (wsB) {
			const valuesB: unknown[][] = [];
			for (let r = 0; r < wsB.cells.maxDataRow + 1; r++) {
				const row: unknown[] = [];
				for (let c = 0; c < wsB.cells.maxDataColumn + 1; c++) {
					row.push(wsB.cells.get(r, c).value);
				}
				valuesB.push(row);
			}
			expect(valuesB).toEqual([[3], [4]]);
		}
	});

	it("can import string values", async () => {
		const handle = await importWorkbook({
			StringSheet: [["hello"]],
		});
		const workbook = handle.workbook;
		const ws = workbook.worksheets.get("StringSheet");
		expect(ws).toBeTruthy();
		if (ws) {
			expect(ws.cells.get(0, 0).value).toBe("hello");
		}
	});

	it("can import number values", async () => {
		const handle = await importWorkbook({
			NumberSheet: [[123]],
		});
		const workbook = handle.workbook;
		const ws = workbook.worksheets.get("NumberSheet");
		expect(ws).toBeTruthy();
		if (ws) {
			expect(ws.cells.get(0, 0).value).toBe(123);
		}
	});

	it("can import boolean values", async () => {
		const handle = await importWorkbook({
			BooleanSheet: [[true, false]],
		});
		const workbook = handle.workbook;
		const ws = workbook.worksheets.get("BooleanSheet");
		expect(ws).toBeTruthy();
		if (ws) {
			expect(ws.cells.get(0, 0).value).toBe(true);
			expect(ws.cells.get(0, 1).value).toBe(false);
		}
	});

	it("can import date values", async () => {
		const testDate = new Date("2023-01-01T12:34:56Z");
		const handle = await importWorkbook({
			DateSheet: [[testDate]],
		});
		const workbook = handle.workbook;
		const ws = workbook.worksheets.get("DateSheet");
		expect(ws).toBeTruthy();
		if (ws) {
			const cell = ws.cells.get(0, 0).value;
			expect(cell instanceof Date || typeof cell === "number").toBe(true);
		}
	});

	it("can import formulas", async () => {
		const handle = await importWorkbook({
			FormulaSheet: [["=SUM(1,2,3)"]],
		});
		const workbook = handle.workbook;
		const ws = workbook.worksheets.get("FormulaSheet");
		expect(ws).toBeTruthy();
		if (ws) {
			const formulaCell = ws.cells.get(0, 0);
			if (formulaCell.isFormula) {
				expect(formulaCell.formula).toBe("=SUM(1,2,3)");
			} else {
				throw new Error("Formula cell not imported as formula");
			}
		}
	});
});
