import { describe, expect, it } from "vitest";
import type { CellRef, Ref } from "../models/Reference.ts";
import { parseCellRef, parseRef, resolveColumnIndex, resolveRowIndex } from "../services/reference";

describe("parseCellRef", () => {
	it("parses cell ref string", () => {
		const [col, row] = parseCellRef("B2");
		expect(col).toBe(1);
		expect(row).toBe(1);
	});

	it("throws on invalid string", () => {
		expect(() => parseCellRef("ZZZ" as CellRef)).toThrow();
	});

	it("throws on invalid array", () => {
		expect(() => parseCellRef(["Sheet1"] as unknown as CellRef)).toThrow();
	});

	it("throws on invalid string", () => {
		expect(() => parseCellRef("Sheet1!ZZZ" as CellRef)).toThrow();
	});
});

describe("parseRef", () => {
	it("can parse column", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef("C");
		expect(colStart).toBe(2);
		expect(rowStart).toBe(null);
		expect(colEnd).toBe(2);
		expect(rowEnd).toBe(null);
	});
	it("can parse row numeric", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef(3);
		expect(colStart).toBe(null);
		expect(rowStart).toBe(2);
		expect(colEnd).toBe(null);
		expect(rowEnd).toBe(2);
	});
	it("can parse row string", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef("3");
		expect(colStart).toBe(null);
		expect(rowStart).toBe(2);
		expect(colEnd).toBe(null);
		expect(rowEnd).toBe(2);
	});
	it("can parse cell", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef("C3");
		expect(colStart).toBe(2);
		expect(rowStart).toBe(2);
		expect(colEnd).toBe(2);
		expect(rowEnd).toBe(2);
	});

	it("can parse cell range", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef("A1:C3");
		expect(colStart).toBe(0);
		expect(rowStart).toBe(0);
		expect(colEnd).toBe(2);
		expect(rowEnd).toBe(2);
	});

	it("can parse cell range array", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef(["A1", "C3"]);
		expect(colStart).toBe(0);
		expect(rowStart).toBe(0);
		expect(colEnd).toBe(2);
		expect(rowEnd).toBe(2);
	});

	it("can parse row range", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef("1:5");
		expect(colStart).toBe(null);
		expect(rowStart).toBe(0);
		expect(colEnd).toBe(null);
		expect(rowEnd).toBe(4);
	});

	it("can parse row range numeric array", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef([1, 5]);
		expect(colStart).toBe(null);
		expect(rowStart).toBe(0);
		expect(colEnd).toBe(null);
		expect(rowEnd).toBe(4);
	});

	it("can parse row range string array", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef(["1", "5"]);
		expect(colStart).toBe(null);
		expect(rowStart).toBe(0);
		expect(colEnd).toBe(null);
		expect(rowEnd).toBe(4);
	});

	it("can parse column range", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef("A:C");
		expect(colStart).toBe(0);
		expect(rowStart).toBe(null);
		expect(colEnd).toBe(2);
		expect(rowEnd).toBe(null);
	});

	it("can parse column range array", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef(["A", "C"]);
		expect(colStart).toBe(0);
		expect(rowStart).toBe(null);
		expect(colEnd).toBe(2);
		expect(rowEnd).toBe(null);
	});

	it("can parse column-row range", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef("A:3");
		expect(colStart).toBe(0);
		expect(rowStart).toBe(null);
		expect(colEnd).toBe(null);
		expect(rowEnd).toBe(2);
	});

	it("can parse column-row range numeric array", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef(["A", 3]);
		expect(colStart).toBe(0);
		expect(rowStart).toBe(null);
		expect(colEnd).toBe(null);
		expect(rowEnd).toBe(2);
	});

	it("can parse column-row range string array", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef(["A", "3"]);
		expect(colStart).toBe(0);
		expect(rowStart).toBe(null);
		expect(colEnd).toBe(null);
		expect(rowEnd).toBe(2);
	});

	it("can parse row-column range", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef("1:C");
		expect(colStart).toBe(null);
		expect(rowStart).toBe(0);
		expect(colEnd).toBe(2);
		expect(rowEnd).toBe(null);
	});

	it("can parse row-column range numeric array", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef([1, "C"]);
		expect(colStart).toBe(null);
		expect(rowStart).toBe(0);
		expect(colEnd).toBe(2);
		expect(rowEnd).toBe(null);
	});

	it("can parse row-column range string array", () => {
		const [colStart, rowStart, colEnd, rowEnd] = parseRef(["1", "C"]);
		expect(colStart).toBe(null);
		expect(rowStart).toBe(0);
		expect(colEnd).toBe(2);
		expect(rowEnd).toBe(null);
	});

	it("throws InvalidArgumentError if range ends before it starts (array)", () => {
		// endCol < startCol
		expect(() => parseRef(["C3", "A1"])).toThrowError(/Range ends before it starts/);
		// endRow < startRow
		expect(() => parseRef(["A3", "A1"])).toThrowError(/Range ends before it starts/);
	});

	it("throws InvalidArgumentError if range ends before it starts (string)", () => {
		// endCol < startCol
		expect(() => parseRef("C3:A1" as unknown as Ref)).toThrowError(/Range ends before it starts/);
		// endRow < startRow
		expect(() => parseRef("A3:A1" as unknown as Ref)).toThrowError(/Range ends before it starts/);
	});
});

describe("resolveColumnIndex", () => {
	it("converts A to 0", () => {
		expect(resolveColumnIndex("A")).toBe(0);
	});
	it("converts Z to 25", () => {
		expect(resolveColumnIndex("Z")).toBe(25);
	});
	// Only single letter columns are supported by ColumnComponent type
});

describe("resolveRowIndex", () => {
	it("parses string row", () => {
		expect(resolveRowIndex("42")).toBe(41);
	});
	it("parses number row", () => {
		expect(resolveRowIndex(7)).toBe(6);
	});
	it("throws on invalid row", () => {
		expect(() => resolveRowIndex("foo" as unknown as number)).toThrow();
	});
});
