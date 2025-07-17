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
import type { Handle } from "../models/Handle.ts";
import writeWorkbook from "./writeWorkbook.ts";

const rows = [
	["A", "B", "C"],
	["D", "E", "F"],
	["G", "H", "I"],
];

describe("writeWorkbook", () => {
	it("can write workbook", async () => {
		const driveRef = getDefaultDriveRef();
		const remoteItemPath = driveItemPath(generateTempFileName("xlsx"));
		let remoteItemRef = await createWorkbook(driveRef, remoteItemPath);

		const workbook = new AsposeCells.Workbook();
		const worksheet = workbook.worksheets.get(0);
		for (let r = 0; r < rows.length; r++) {
			for (let c = 0; c < rows[r].length; c++) {
				worksheet.cells.get(r, c).putValue(rows[r][c]);
			}
		}
		const handle: Handle = {
			workbook,
			remoteItemRef,
		};

		remoteItemRef = await writeWorkbook(handle);

		const remoteWorksheetRef = await getWorkbookWorksheetByName(remoteItemRef, defaultWorkbookWorksheetName);
		const usedRange = await getWorkbookWorksheetUsedRange(remoteWorksheetRef);
		expect(usedRange.values).toEqual(rows);

		await deleteDriveItem(remoteItemRef);
	});
});
