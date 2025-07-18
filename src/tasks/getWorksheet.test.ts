import { describe, expect, it } from "vitest";
import type { WorksheetName } from "../models/Worksheet";
import createWorkbook from "./createWorkbook";
import getWorksheet from "./getWorksheet";

describe("getWorksheet", () => {
	it("returns worksheet by exact name", async () => {
		const wb = await createWorkbook({
			Sheet1: [[1]],
			Sheet2: [[2]],
		});
		const ws1 = getWorksheet(wb, "Sheet1" as WorksheetName);
		expect(ws1.name).toBe("Sheet1");
		const ws2 = getWorksheet(wb, "Sheet2" as WorksheetName);
		expect(ws2.name).toBe("Sheet2");
	});

	it("throws if worksheet does not exist", async () => {
		const wb = await createWorkbook({
			Sheet1: [[1]],
		});
		expect(() => getWorksheet(wb, "NoSuchSheet" as WorksheetName)).toThrow();
	});

	it("is case-insensitive", async () => {
		const wb = await createWorkbook({
			SheetA: [[1]],
		});
		const ws = getWorksheet(wb, "sheeta" as WorksheetName);
		expect(ws.name).toBe("SheetA");
	});
});
