import ExcelJS from "exceljs";
import { PassThrough } from "node:stream";
import { describe, expect, it } from "vitest";
import type { WorksheetName } from "../models/Worksheet";
import { csvToExcel } from "../services/csvToExcel";

import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import { unlink } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { LocalFilePath } from "../models/LocalFilePath";

function streamFromString(str: string) {
	const stream = new PassThrough();
	stream.end(str);
	return stream;
}

describe("csvToExcel", () => {
	it("can convert simple CSV to Excel", async () => {
		const file = join(tmpdir(), generateTempFileName("xlsx")) as LocalFilePath;
		const csv = "A,B,C\n1,2,3\n4,5,6";
		const stream = streamFromString(csv);
		await csvToExcel(stream, file);

		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.readFile(file);
		const worksheet = workbook.getWorksheet(1);
		expect(worksheet?.name).toBe("Sheet1");
		expect(Array.isArray(worksheet?.getRow(1).values)).toBe(true);
		expect((worksheet?.getRow(1).values as (string | number | null)[]).slice(1)).toEqual(["A", "B", "C"]);
		expect((worksheet?.getRow(2).values as (string | number | null)[]).slice(1)).toEqual(["1", "2", "3"]);
		expect((worksheet?.getRow(3).values as (string | number | null)[]).slice(1)).toEqual(["4", "5", "6"]);

		await unlink(file);
	});

	it("can escape special XML characters in cell values", async () => {
		const file = join(tmpdir(), generateTempFileName("xlsx")) as LocalFilePath;
		const csv = 'A,B\n<foo>,"bar & baz"';
		await csvToExcel(streamFromString(csv), file);

		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.readFile(file);
		const ws = workbook.getWorksheet(1);
		expect(Array.isArray(ws?.getRow(2).values)).toBe(true);
		expect((ws?.getRow(2).values as (string | number | null)[]).slice(1)).toEqual(["<foo>", "bar & baz"]);

		await unlink(file);
	});

	it("can handle empty cells and rows", async () => {
		const file = join(tmpdir(), generateTempFileName("xlsx")) as LocalFilePath;
		const csv = "A,B,C\n,,\n1,,3";
		await csvToExcel(streamFromString(csv), file);

		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.readFile(file);
		const ws = workbook.getWorksheet(1);
		expect(Array.isArray(ws?.getRow(2).values)).toBe(true);
		expect((ws?.getRow(1).values as (string | number | null)[]).slice(1)).toEqual(["A", "B", "C"]);
		expect((ws?.getRow(2).values as (string | number | null)[]).slice(1)).toEqual([]);
		expect((ws?.getRow(3).values as (string | number | null)[]).slice(1)).toEqual(["1", undefined, "3"]);

		await unlink(file);
	});

	it("can use the provided worksheet name", async () => {
		const file = join(tmpdir(), generateTempFileName("xlsx")) as LocalFilePath;
		const csv = "A\nB";
		const worksheetName = "MySheet!" as WorksheetName;
		await csvToExcel(streamFromString(csv), file, { worksheetName });

		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.readFile(file);
		expect(workbook.worksheets[0].name).toBe(worksheetName);

		await unlink(file);
	});
});
