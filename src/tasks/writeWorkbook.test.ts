import createWorkbook from "microsoft-graph/createWorkbook";
import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import tryDeleteDriveItem from "microsoft-graph/tryDeleteDriveItem";
import { createWorkbookRangeRef } from "microsoft-graph/workbookRange";
import { createDefaultWorkbookWorksheetRef } from "microsoft-graph/workbookWorksheet";
import writeWorkbookRows from "microsoft-graph/writeWorkbookRows";
import { describe, it } from "vitest";
import readWorkbook from "./readWorkbook.ts";
import writeWorkbook from "./writeWorkbook.ts";

const rows = [
	[{ value: "A" }, { value: "B" }, { value: "C" }],
	[{ value: "D" }, { value: "E" }, { value: "F" }],
	[{ value: "G" }, { value: "H" }, { value: "I" }],
];

describe("writeWorkbook", () => {
	it("creates small XLSX file", async () => {
		const workbookName = generateTempFileName("xlsx");
		const workbookPath = driveItemPath(workbookName);
		const driveRef = getDefaultDriveRef();
		const item = await createWorkbook(driveRef, workbookPath);
		const worksheetRef = createDefaultWorkbookWorksheetRef(item);
		const rangeRef = createWorkbookRangeRef(worksheetRef, "A1");
		await writeWorkbookRows(rangeRef, rows);

		const hdl = await readWorkbook(item);

		await writeWorkbook(hdl);

		await tryDeleteDriveItem(item);
	});
});
