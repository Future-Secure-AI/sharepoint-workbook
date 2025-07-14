import { getDefaultDriveRef } from "microsoft-graph/drive";
import { describe, expect, it } from "vitest";
import importWorkbook from "./importWorkbook";
import writeWorkbook from "./writeWorkbook.ts";
import type { DriveItemId } from "microsoft-graph/DriveItem";

function getSmallSet() {
	return [
		[{ value: "A" }, { value: "B" }, { value: "C" }],
		[{ value: "D" }, { value: "E" }, { value: "F" }],
		[{ value: "G" }, { value: "H" }, { value: "I" }],
	];
}

describe("writeWorkbook", { timeout: 15 * 60 * 1000 }, () => {
	it("creates small XLSX file", async () => {
		const hdl = await importWorkbook([
			{ name: "Sheet1", rows: getSmallSet() },
			{ name: "Sheet2", rows: getSmallSet() },
		]);
		const driveRef = getDefaultDriveRef();
		const itemId = "dummy-item-id" as unknown as DriveItemId;
		const itemRef = { ...driveRef, itemId };
		await expect(writeWorkbook({ ...hdl, itemRef })).resolves.toBeUndefined();
	});

	it("throws for unsupported file extension", async () => {
		const worksheets = [{ name: "Sheet1", rows: getSmallSet() }];
		const handle = await importWorkbook(worksheets);
		const driveRef = getDefaultDriveRef();
		const itemId = "dummy-item-id" as unknown as DriveItemId;
		const itemRef = { ...driveRef, itemId };
		await expect(writeWorkbook({ ...handle, itemRef })).rejects.toThrow(/Unsupported file extension/);
	});
});
