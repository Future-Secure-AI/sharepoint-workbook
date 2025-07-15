import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import tryDeleteDriveItem from "microsoft-graph/tryDeleteDriveItem";
import { describe, expect, it } from "vitest";
import importWorkbook from "./importWorkbook";
import writeWorkbookByPath from "./writeWorkbookByPath";

const rows = [
	["A", "B", "C"],
	["D", "E", "F"],
	["G", "H", "I"],
];

describe("writeWorkbookByPath", () => {
	it("writes a small XLSX file to a path", async () => {
		const driveRef = getDefaultDriveRef();
		const itemPath = driveItemPath(generateTempFileName("xlsx"));

		const handle = await importWorkbook([
			{ name: "Sheet1", rows: rows },
			{ name: "Sheet2", rows: rows },
		]);

		const item = await writeWorkbookByPath(handle, driveRef, itemPath);

		await tryDeleteDriveItem(item);
	});

	it("throws for unsupported file extension", async () => {
		const driveRef = getDefaultDriveRef();
		const itemPath = driveItemPath(generateTempFileName("txt"));

		const handle = await importWorkbook([{ name: "Sheet1", rows: rows }]);
		await expect(writeWorkbookByPath(handle, driveRef, itemPath)).rejects.toThrow(/Unsupported file extension/);
	});
});
