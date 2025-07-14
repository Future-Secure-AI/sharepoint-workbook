import { describe, expect, it } from "vitest";
import importWorkbook from "./importWorkbook";
import writeWorkbookByPath from "./writeWorkbookByPath";

import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import tryDeleteDriveItem from "microsoft-graph/tryDeleteDriveItem";

function getSmallSet() {
	return [
		[{ value: "A" }, { value: "B" }, { value: "C" }],
		[{ value: "D" }, { value: "E" }, { value: "F" }],
		[{ value: "G" }, { value: "H" }, { value: "I" }],
	];
}

describe("writeWorkbookByPath", () => {
	it("writes a small XLSX file to a path", async () => {
		const driveRef = getDefaultDriveRef();
		const itemPath = driveItemPath(generateTempFileName("xlsx"));

		const handle = await importWorkbook([
			{ name: "Sheet1", rows: getSmallSet() },
			{ name: "Sheet2", rows: getSmallSet() },
		]);

		const item = await writeWorkbookByPath(handle, driveRef, itemPath);

		await tryDeleteDriveItem(item);
	});

	it("throws for unsupported file extension", async () => {
		const driveRef = getDefaultDriveRef();
		const itemPath = driveItemPath(generateTempFileName("txt"));

		const handle = await importWorkbook([{ name: "Sheet1", rows: getSmallSet() }]);
		await expect(writeWorkbookByPath(handle, driveRef, itemPath)).rejects.toThrow(/Unsupported file extension/);
	});
});
