import AsposeCells from "aspose.cells.node";
import createWorkbook from "microsoft-graph/createWorkbook";
import deleteDriveItem from "microsoft-graph/deleteDriveItem";
import getWorkbookWorksheetByName from "microsoft-graph/dist/cjs/operations/workbookWorksheet/getWorkbookWorksheetByName";
import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import getWorkbookWorksheetUsedRange from "microsoft-graph/getWorkbookWorksheetUsedRange";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import { defaultWorkbookWorksheetName } from "microsoft-graph/workbookWorksheet";
import { describe, expect, it } from "vitest";
import type { Workbook } from "../models/Workbook";
import saveWorkbook from "./saveWorkbook.ts";

const rows = [
	["A", "B", "C"],
	["D", "E", "F"],
	["G", "H", "I"],
];

describe("writeWorkbook", () => {
	it("can save workbook", async () => {
		const driveRef = getDefaultDriveRef();
		const remoteItemPath = driveItemPath(generateTempFileName("xlsx"));
		let remoteItem = await createWorkbook(driveRef, remoteItemPath);

		const workbook = new AsposeCells.Workbook() as Workbook;
		const worksheet = workbook.worksheets.get(0);
		for (let r = 0; r < rows.length; r++) {
			for (let c = 0; c < rows[r].length; c++) {
				worksheet.cells.get(r, c).putValue(rows[r][c]);
			}
		}
		workbook.remoteItem = remoteItem;

		remoteItem = await saveWorkbook(workbook);

		const remoteWorksheetRef = await getWorkbookWorksheetByName(remoteItem, defaultWorkbookWorksheetName);
		const usedRange = await getWorkbookWorksheetUsedRange(remoteWorksheetRef);
		expect(usedRange.values).toEqual(rows);

		await deleteDriveItem(remoteItem);
	});

	it("can not save if not savedAs before", async () => {
		const workbook = new AsposeCells.Workbook() as Workbook;
		await expect(saveWorkbook(workbook)).rejects.toThrow("Workbook not over-writable. Use `saveWorkbookAs` instead.");
	});
});
