import { describe, expect, it } from "vitest";
import type { Handle } from "../models/Handle";
import type { WorksheetWrite } from "../models/Worksheet";
import importWorkbook from "./importWorkbook";
import optimizeWorkbook from "./optimizeWorkbook";

describe("optimizeWorkbook", () => {
	it("returns a valid ratio for a real workbook", async () => {
		const worksheets: WorksheetWrite[] = [
			{
				name: "Sheet1",
				rows: [
					[1, 2, 3],
					[4, 5, 6],
				],
			},
		];
		const hdl = await importWorkbook(worksheets);
		const ratio = await optimizeWorkbook(hdl, { compressionLevel: 6 });
		expect(ratio).toBeGreaterThan(0);
		expect(ratio).toBeLessThanOrEqual(1);
	});

	it("throws if file does not exist", async () => {
		await expect(optimizeWorkbook({ id: "bad-id" } as Handle)).rejects.toThrow();
	});
});
