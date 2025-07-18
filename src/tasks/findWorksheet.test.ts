import { describe, expect, it } from "vitest";
import createWorkbook from "./createWorkbook";
import findWorksheet from "./findWorksheet";

describe("findWorksheet", () => {
	it("finds worksheet by exact name (case-insensitive)", async () => {
		const wb = await createWorkbook({
			SheetA: [[1]],
			SheetB: [[2]],
		});
		const ws = findWorksheet(wb, "sheeta");
		expect(ws.name).toBe("SheetA");
		const ws2 = findWorksheet(wb, "SHEETB");
		expect(ws2.name).toBe("SheetB");
	});

	it("finds worksheet by glob pattern", async () => {
		const wb = await createWorkbook({
			Data2023: [[1]],
			Data2024: [[2]],
			Other: [[3]],
		});
		const ws = findWorksheet(wb, "Data202*");
		expect(["Data2023", "Data2024"]).toContain(ws.name);
	});

	it("throws if no worksheet matches", async () => {
		const wb = await createWorkbook({
			Sheet1: [[1]],
		});
		expect(() => findWorksheet(wb, "NoMatch")).toThrow();
	});

	it("returns first match if multiple match", async () => {
		const wb = await createWorkbook({
			Alpha: [[1]],
			Alphanumeric: [[2]],
			Beta: [[3]],
		});
		const ws = findWorksheet(wb, "Alph*");
		expect(["Alpha", "Alphanumeric"]).toContain(ws.name);
		// Should be the first match in order
		expect(ws.name).toBe("Alpha");
	});
});
