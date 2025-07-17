import { describe, expect, it } from "vitest";
import type { CellRef, RangeRef } from "../models/Reference.ts";
import { columnComponentToNumber, parseCellReference, parseRangeReference, rowComponentToNumber } from "../services/reference";

describe("parseCellReference", () => {
	it("parses cell ref string", () => {
		const [col, row] = parseCellReference("B2");
		expect(col).toBe(2);
		expect(row).toBe(2);
	});

	it("throws on invalid string", () => {
		expect(() => parseCellReference("ZZZ" as CellRef)).toThrow();
	});

	it("throws on invalid array", () => {
		expect(() => parseCellReference(["Sheet1"] as unknown as CellRef)).toThrow();
	});

	it("throws on invalid string", () => {
		expect(() => parseCellReference("Sheet1!ZZZ" as CellRef)).toThrow();
	});
});

describe("parseRangeReference", () => {
	it("parses array range ref", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRangeReference(["A1", "C3"]);
		expect(colStart).toBe(1);
		expect(rowStart).toBe(1);
		expect(colEnd).toBe(3);
		expect(rowEnd).toBe(3);
	});

	it("parses string range ref", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRangeReference("A1:C3");
		expect(colStart).toBe(1);
		expect(rowStart).toBe(1);
		expect(colEnd).toBe(3);
		expect(rowEnd).toBe(3);
	});

	it("throws on invalid array", () => {
		expect(() => parseRangeReference(["A1"] as unknown as RangeRef)).toThrow();
	});

	it("throws on invalid string", () => {
		expect(() => parseRangeReference("A1" as unknown as RangeRef)).toThrow();
	});

	it("throws InvalidArgumentError if range ends before it starts (array)", () => {
		// endCol < startCol
		expect(() => parseRangeReference(["C3", "A1"])).toThrowError(/Range ends before it starts/);
		// endRow < startRow
		expect(() => parseRangeReference(["A3", "A1"])).toThrowError(/Range ends before it starts/);
	});

	it("throws InvalidArgumentError if range ends before it starts (string)", () => {
		// endCol < startCol
		expect(() => parseRangeReference("C3:A1" as unknown as RangeRef)).toThrowError(/Range ends before it starts/);
		// endRow < startRow
		expect(() => parseRangeReference("A3:A1" as unknown as RangeRef)).toThrowError(/Range ends before it starts/);
	});
});

describe("columnComponentToNumber", () => {
	it("converts A to 1", () => {
		expect(columnComponentToNumber("A")).toBe(1);
	});
	it("converts Z to 26", () => {
		expect(columnComponentToNumber("Z")).toBe(26);
	});
	// Only single letter columns are supported by ColumnComponent type
});

describe("rowComponentToNumber", () => {
	it("parses string row", () => {
		expect(rowComponentToNumber("42")).toBe(42);
	});
	it("parses number row", () => {
		expect(rowComponentToNumber(7)).toBe(7);
	});
	it("throws on invalid row", () => {
		expect(() => rowComponentToNumber("foo" as unknown as number)).toThrow();
	});
});
