import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import { describe, expect, it } from "vitest";
import type { CellWrite } from "../models/Cell";
import { normalizeCellWrite } from "../services/cell";

describe("normalizeCellWrite", () => {
	it("should wrap string value in CellWrite", () => {
		expect(normalizeCellWrite("foo")).toEqual({ value: "foo" });
	});

	it("should wrap number value in CellWrite", () => {
		expect(normalizeCellWrite(42)).toEqual({ value: 42 });
	});

	it("should wrap boolean value in CellWrite", () => {
		expect(normalizeCellWrite(true)).toEqual({ value: true });
		expect(normalizeCellWrite(false)).toEqual({ value: false });
	});

	it("should wrap Date value in CellWrite", () => {
		const date = new Date();
		expect(normalizeCellWrite(date)).toEqual({ value: date });
	});

	it("should return CellWrite as-is if already an object", () => {
		const obj: CellWrite = { value: "bar", format: "bold" } as unknown as CellWrite;
		expect(normalizeCellWrite(obj)).toBe(obj);
	});

	it("should throw InvalidArgumentError for unsupported types", () => {
		expect(() => normalizeCellWrite(undefined as unknown as CellWrite)).toThrow(InvalidArgumentError);
		expect(() => normalizeCellWrite(Symbol("x") as unknown as CellWrite)).toThrow(InvalidArgumentError);
	});
});
