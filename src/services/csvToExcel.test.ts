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
	it("uses shared strings for short values and inline for long values", async () => {
		// Prepare CSV with repeated short and one long value
		const csvRows = [
			["foo", "bar", "baz"],
			["foo", "bar", "baz"],
			["longstringvalue", "bar", "baz"],
			["foo", "bar", "baz"],
		];
		const csv = csvRows.map((r) => r.join(",")).join("\n");
		const file = join(tmpdir(), generateTempFileName("xlsx")) as LocalFilePath;
		await csvToExcel(streamFromString(csv), file);

		// Use yauzl to read the zip file
		const yauzl = await import("yauzl");
		const getFileContent = (zipPath: string): Promise<string> => {
			return new Promise((resolve, reject) => {
				yauzl.open(file, { lazyEntries: true }, (err, zipfile) => {
					if (err || !zipfile) return reject(err);
					let found = false;
					zipfile.readEntry();
					zipfile.on("entry", (entry) => {
						if (entry.fileName === zipPath) {
							found = true;
							zipfile.openReadStream(entry, (err, readStream) => {
								if (err || !readStream) return reject(err);
								let data = "";
								readStream.on("data", (chunk) => {
									data += chunk;
								});
								readStream.on("end", () => resolve(data));
							});
						} else {
							zipfile.readEntry();
						}
					});
					zipfile.on("end", () => {
						if (!found) reject(new Error(`${zipPath} not found`));
					});
				});
			});
		};
		const sharedStringsXml = await getFileContent("xl/sharedStrings.xml");
		expect(sharedStringsXml).toContain("foo");
		expect(sharedStringsXml).toContain("bar");
		expect(sharedStringsXml).toContain("baz");
		expect(sharedStringsXml).not.toContain("longstringvalue");

		const sheetXml = await getFileContent("xl/worksheets/sheet1.xml");
		expect(sheetXml).toContain('<c r="A3" t="inlineStr"><is><t>longstringvalue</t></is></c>');
		expect(sheetXml).toContain('<c r="A1" t="s"><v>0</v></c>');
		await unlink(file);
	});
});
