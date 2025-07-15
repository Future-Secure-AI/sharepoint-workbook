import ExcelJS from "exceljs";
import createDriveItemContent from "microsoft-graph/createDriveItemContent";
import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import tryDeleteDriveItem from "microsoft-graph/tryDeleteDriveItem";
import { Readable } from "node:stream";
import { describe, expect, it } from "vitest";
import { getLatestRevisionFilePath } from "../services/workingFolder";
import importWorkbook from "./importWorkbook";
import readWorkbook from "./readWorkbook";
import writeWorkbookByPath from "./writeWorkbookByPath";

const rows = [
	["A", "B", "C"],
	["D", "E", "F"],
	["G", "H", "I"],
];

describe("readWorkbook", () => {
	it("can read XLSX workbook", async () => {
		const driveRef = getDefaultDriveRef();
		const itemPath = driveItemPath(generateTempFileName("xlsx"));
		const writeHandle = await importWorkbook([
			{ name: "Sheet1", rows: rows },
			{ name: "Sheet2", rows: rows },
		]);
		const item = await writeWorkbookByPath(writeHandle, driveRef, itemPath);
		const readHandle = await readWorkbook(item);

		const file = await getLatestRevisionFilePath(readHandle.id);
		const wb = new ExcelJS.Workbook();
		await wb.xlsx.readFile(file);
		["Sheet1", "Sheet2"].forEach((sheet) => {
			const ws = wb.getWorksheet(sheet);
			expect(ws).toBeTruthy();
			if (ws) {
				expect(
					ws
						.getSheetValues()
						.slice(1)
						.map((row) => (Array.isArray(row) ? row.slice(1) : [])),
				).toEqual(rows);
			}
		});
		await tryDeleteDriveItem(item);
	});

	it("can read CSV workbook", async () => {
		const driveRef = getDefaultDriveRef();
		const itemPath = driveItemPath(generateTempFileName("csv"));

		const text = `A,B,C\nD,E,F\nG,H,I`;
		const length = Buffer.byteLength(text, "utf8");
		const stream = Readable.from([text]);
		const item = await createDriveItemContent(driveRef, itemPath, stream, length);

		const readHandle = await readWorkbook(item);
		const file = await getLatestRevisionFilePath(readHandle.id);
		const workbook = new ExcelJS.Workbook();
		await workbook.xlsx.readFile(file);
		const worksheet = workbook.getWorksheet(1); // 1-based index
		expect(worksheet).toBeTruthy();

		if (worksheet) {
			const values = worksheet
				.getSheetValues()
				.filter((row) => Array.isArray(row)) // filter out null/undefined
				.map((row) => row.slice(1)); // remove first empty column
			expect(values).toEqual(rows);
		}
		await tryDeleteDriveItem(item);
	});
});
