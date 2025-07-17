import closeWorkbookSession from "microsoft-graph/closeWorkbookSession";
import createDriveItemContent from "microsoft-graph/createDriveItemContent";
import createWorkbookAndStartSession from "microsoft-graph/createWorkbookAndStartSession";
import deleteDriveItem from "microsoft-graph/deleteDriveItem";
import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { sleep } from "microsoft-graph/sleep";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import updateWorkbookRange from "microsoft-graph/updateWorkbookRange";
import { createWorkbookRangeRef } from "microsoft-graph/workbookRange";
import { createWorkbookWorksheetRef, defaultWorkbookWorksheetId, defaultWorkbookWorksheetName } from "microsoft-graph/workbookWorksheet";
import { Readable } from "node:stream";
import { describe, expect, it } from "vitest";
import openWorkbook from "./openWorkbook";

const rows = [
	["A", "B", "C"],
	["D", "E", "F"],
	["G", "H", "I"],
];

describe("openWorkbook", () => {
	it("reads XLSX workbooks", async () => {
		const workbook = await createWorkbookAndStartSession(getDefaultDriveRef(), driveItemPath(generateTempFileName("xlsx")));
		await updateWorkbookRange(createWorkbookRangeRef(createWorkbookWorksheetRef(workbook, defaultWorkbookWorksheetId), "A1:C3"), { values: rows });
		await closeWorkbookSession(workbook);
		await sleep(1000);

		const { workbook: wb } = await openWorkbook(workbook);

		const ws = wb.worksheets.get(0);
		expect(ws).toBeTruthy();
		expect(ws.name).toBe(defaultWorkbookWorksheetName);
		const values = Array.from({ length: ws.cells.maxDataRow + 1 }, (_, r) => Array.from({ length: ws.cells.maxDataColumn + 1 }, (_, c) => ws.cells.get(r, c).value));
		expect(values).toEqual(rows);

		await deleteDriveItem(workbook);
	});

	it("reads CSV workbooks", async () => {
		const text = `A,B,C\nD,E,F\nG,H,I`;
		const item = await createDriveItemContent(getDefaultDriveRef(), driveItemPath(generateTempFileName("csv")), Readable.from([text]), Buffer.byteLength(text, "utf8"));

		const { workbook: wb } = await openWorkbook(item);

		const ws = wb.worksheets.get(0);
		expect(ws).toBeTruthy();
		const values = Array.from({ length: ws.cells.maxDataRow + 1 }, (_, r) => Array.from({ length: ws.cells.maxDataColumn + 1 }, (_, c) => ws.cells.get(r, c).value));
		expect(values).toEqual(rows);

		await deleteDriveItem(item);
	});
});
