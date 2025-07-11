import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import tryDeleteDriveItem from "microsoft-graph/tryDeleteDriveItem";
import { describe, expect, it } from "vitest";
import writeWorkbook from "./writeWorkbook.ts";

function getSmallSet() {
	return [
		[{ value: "A" }, { value: "B" }, { value: "C" }],
		[{ value: "D" }, { value: "E" }, { value: "F" }],
		[{ value: "G" }, { value: "H" }, { value: "I" }],
	];
}

describe("writeWorkbook", { timeout: 15 * 60 * 1000 }, () => {
	it("creates small XLSX file", async () => {
		const rows = getSmallSet();
		const itemName = generateTempFileName("xlsx");
		const itemPath = driveItemPath(itemName);
		const driveRef = getDefaultDriveRef();
		const worksheets = [
			{ name: "Sheet1", rows },
			{ name: "Sheet2", rows },
		];
		const item = await writeWorkbook(driveRef, itemPath, worksheets);
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
		const worksheets = [{ name: "Sheet1", rows }];
		await expect(writeWorkbook(driveRef, itemPath, worksheets)).rejects.toThrow(/Unsupported file extension/);
	});
});
