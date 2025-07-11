/**
 * Create a workbook.
 * @module createWorkbook
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import ExcelJS from "exceljs";
import InvalidArgumentError from "microsoft-graph/dist/cjs/errors/InvalidArgumentError";
import type { Cell } from "microsoft-graph/dist/cjs/models/Cell";
import type { DriveRef } from "microsoft-graph/dist/cjs/models/Drive";
import type { DriveItemPath, DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";
import type { WorkbookWorksheetName } from "microsoft-graph/dist/cjs/models/WorkbookWorksheet";
import createDriveItemContent from "microsoft-graph/dist/cjs/operations/driveItem/createDriveItemContent";
import { getEnvironmentVariable } from "microsoft-graph/environmentVariable";
import { randomUUID } from "node:crypto";
import { createReadStream, createWriteStream, promises as fs } from "node:fs";
import { tmpdir } from "node:os";
import { extname, join as pathJoin } from "node:path";
import { appendRow } from "../services/excelJs.ts";

/**
 * Options for creating a new workbook file.
 * @property {WorkbookWorksheetName} [sheetName] Name of the worksheet to create.
 * @property {"fail" | "replace" | "rename"} [ifAlreadyExists] How to resolve if the file already exists.
 * @property {number} [maxChunkSize] Maximum chunk size for upload (in bytes).
 * @property {(preparedCount: number, writtenCount: number, preparedPerSecond: number, writtenPerSecond: number) =&lt; void} [progress] Progress callback.
 * @property {string} [workingFolder] Working folder for temporary file storage. Defaults to the `WORKING_FOLDER` env, then the OS temporary folder if not set.
 */
export type CreateOptions = {
	ifAlreadyExists?: "fail" | "replace" | "rename";
	maxChunkSize?: number;
	progress?: (preparedCount: number, writtenCount: number, preparedPerSecond: number, writtenPerSecond: number) => void;
	workingFolder?: string;
};

/**
 * Creates a new workbook (.xlsx) in the specified parent location with the provided rows for multiple sheets.
 * @param {DriveRef | DriveItemRef} parentRef Reference to the parent drive or item where the file will be created.
 * @param {DriveItemPath} itemPath Path (including filename and extension) for the new workbook.
 * @param {Record<WorkbookWorksheetName, Iterable<Partial<Cell>[]> | AsyncIterable<Partial<Cell>[]>>} sheets Object where each key is a sheet name (WorkbookWorksheetName) and the value is an iterable or async iterable of row arrays.
 * @param {CreateOptions} [options] Options for conflict resolution, etc.
 * @returns {Promise<DriveItem & DriveItemRef>} Created DriveItem with reference.
 * @throws {InvalidArgumentError} If the file extension is not supported.
 * @experimental
 */
export default async function writeWorkbook(parentRef: DriveRef | DriveItemRef, itemPath: DriveItemPath, sheets: Record<WorkbookWorksheetName, Iterable<Partial<Cell>[]> | AsyncIterable<Partial<Cell>[]>>, options: CreateOptions = {}): Promise<DriveItem & DriveItemRef> {
	const extension = extname(itemPath);
	if (extension !== ".xlsx") {
		throw new InvalidArgumentError(`Unsupported file extension: ${extension}. Only .xlsx files are supported for workbook creation.`);
	}

	const { ifAlreadyExists: conflictBehavior = "fail", maxChunkSize = 60 * 1024 * 1024, progress = () => {}, workingFolder = getEnvironmentVariable("WORKING_FOLDER", tmpdir()) as string } = options;

	const localFilePath = pathJoin(workingFolder, `${randomUUID()}${extension}`);

	let preparedCells = 0;
	let writtenCells = 0;

	let lastTime = 0;
	let lastPreparedCells = 0;
	let lastWrittenCells = 0;
	try {
		const fileStream = createWriteStream(localFilePath);
		const xls = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: fileStream });

		for (const [sheetName, sheetRows] of Object.entries(sheets)) {
			const worksheet = xls.addWorksheet(sheetName);
			for await (const row of sheetRows) {
				appendRow(worksheet, row);
				preparedCells += row.length;
				progressUpdated();
			}
			worksheet.commit();
		}

		await xls.commit();
		progressUpdated(true);

		const { size } = await fs.stat(localFilePath);
		const stream = createReadStream(localFilePath, { highWaterMark: 1024 * 1024 });
		const item = await createDriveItemContent(parentRef, itemPath, stream, size, {
			conflictBehavior,
			maxChunkSize,
			progress: (bytes) => {
				writtenCells = Math.ceil((bytes / size) * preparedCells);
				progressUpdated();
			},
		});
		progressUpdated(true);
		return item;
	} finally {
		await fs.unlink(localFilePath).catch(() => {});
	}

	function progressUpdated(force: boolean = false): void {
		const time = Date.now();
		const timeDiff = time - lastTime;
		if (force || timeDiff > 1000) {
			const preparedPerSecond = timeDiff ? Math.ceil((preparedCells - lastPreparedCells) / (timeDiff / 1000)) : 0;
			const writtenPerSecond = timeDiff ? Math.ceil((writtenCells - lastWrittenCells) / (timeDiff / 1000)) : 0;
			lastPreparedCells = preparedCells;
			lastWrittenCells = writtenCells;
			lastTime = time;

			progress(preparedCells, writtenCells, preparedPerSecond, writtenPerSecond);
		}
	}
}
