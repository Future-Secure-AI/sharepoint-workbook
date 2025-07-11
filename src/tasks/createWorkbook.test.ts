import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import tryDeleteDriveItem from "microsoft-graph/tryDeleteDriveItem";
import type { WorkbookWorksheetName } from "microsoft-graph/WorkbookWorksheet";
import { describe, expect, it } from "vitest";
import createWorkbook from "./createWorkbook.ts";

function getSmallSet() {
	return [
		[{ value: "A" }, { value: "B" }, { value: "C" }],
		[{ value: "D" }, { value: "E" }, { value: "F" }],
		[{ value: "G" }, { value: "H" }, { value: "I" }],
	];
}

describe("createWorkbook", { timeout: 15 * 60 * 1000 }, () => {
	it("creates small XLSX file", async () => {
		const rows = getSmallSet();
		const itemName = generateTempFileName("xlsx");
		const itemPath = driveItemPath(itemName);
		const driveRef = getDefaultDriveRef();
		const item = await createWorkbook(driveRef, itemPath, {
			["Sheet1" as WorkbookWorksheetName]: rows,
			["Sheet2" as WorkbookWorksheetName]: rows,
		});
		expect(item).toBeTruthy();
		expect(item.name).toBe(itemName);
		expect(item.size).toBeGreaterThan(0);
		await tryDeleteDriveItem(item);
	});

	it("throws for unsupported file extension", async () => {
		const rows = getSmallSet();
		const itemName = generateTempFileName("txt");
		const itemPath = driveItemPath(itemName);
		const driveRef = getDefaultDriveRef();
		await expect(createWorkbook(driveRef, itemPath, { ["Sheet1" as WorkbookWorksheetName]: rows })).rejects.toThrow(/Unsupported file extension/);
	});
});
