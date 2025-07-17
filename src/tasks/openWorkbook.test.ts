import closeWorkbookSession from "microsoft-graph/closeWorkbookSession";
import createWorkbookAndStartSession from "microsoft-graph/createWorkbookAndStartSession";
import deleteDriveItem from "microsoft-graph/deleteDriveItem";
import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { sleep } from "microsoft-graph/sleep";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import updateWorkbookRange from "microsoft-graph/updateWorkbookRange";
import { createWorkbookRangeRef } from "microsoft-graph/workbookRange";
import { createWorkbookWorksheetRef, defaultWorkbookWorksheetId, defaultWorkbookWorksheetName } from "microsoft-graph/workbookWorksheet";
import { describe, expect, it } from "vitest";
import openWorkbook from "./openWorkbook";

const rows = [
	["A", "B", "C"],
	["D", "E", "F"],
	["G", "H", "I"],
];

describe("openWorkbook", () => {
	it("can read workbook by path", async () => {
		const driveRef = getDefaultDriveRef();
		const itemPath = driveItemPath(generateTempFileName("xlsx"));
		const workbook = await createWorkbookAndStartSession(driveRef, itemPath);
		await updateWorkbookRange(createWorkbookRangeRef(createWorkbookWorksheetRef(workbook, defaultWorkbookWorksheetId), "A1:C3"), { values: rows });
		await closeWorkbookSession(workbook);
		await sleep(1000);

		const readHandle = await openWorkbook(driveRef, itemPath);

		const wb = readHandle.workbook;
		const ws = wb.worksheets.get(0);
		expect(ws.name).toBe(defaultWorkbookWorksheetName);
		expect(ws).toBeTruthy();
		const values = Array.from({ length: ws.cells.maxDataRow + 1 }, (_, r) => Array.from({ length: ws.cells.maxDataColumn + 1 }, (_, c) => ws.cells.get(r, c).value));
		expect(values).toEqual(rows);

		await deleteDriveItem(workbook);
	});

	it("throws for missing file", async () => {
		const driveRef = getDefaultDriveRef();
		const itemPath = driveItemPath(generateTempFileName("xlsx"));
		await expect(openWorkbook(driveRef, itemPath)).rejects.toThrow();
	});
});
