/**
 * Copy a drive item.
 * @module createWorkbook
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import ExcelJS from "exceljs";
import type { Cell } from "microsoft-graph/Cell";
import createDriveItemContent from "microsoft-graph/createDriveItemContent";
import type { DriveRef } from "microsoft-graph/Drive";
import type { DriveItemPath, DriveItemRef } from "microsoft-graph/DriveItem";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { WorkbookWorksheetName } from "microsoft-graph/WorkbookWorksheet";
import { randomUUID } from "node:crypto";
import { createReadStream, createWriteStream, promises as fs } from "node:fs";
import { tmpdir } from "node:os";
import { extname, join as pathJoin } from "node:path";
import { appendRow } from "../services/excelJs.ts";

/**
 * Options for creating a new workbook file.
 * @property {WorkbookWorksheetName} [sheetName] Name of the worksheet to create.
 * @property {"fail" | "replace" | "rename"} [conflictResolution] How to resolve name conflicts when uploading the file.
 */
export type CreateOptions = {
	conflictBehavior?: "fail" | "replace" | "rename";
	maxChunkSize?: number;
	progress?: (preparedCount: number, writtenCount: number, preparedPerSecond: number, writtenPerSecond: number) => void;
};

/**
 * Creates a new workbook (.xlsx) in the specified parent location with the provided rows for multiple sheets.
 * @param {DriveRef | DriveItemRef} parentRef Reference to the parent drive or item where the file will be created.
 * @param {DriveItemPath} itemPath Path (including filename and extension) for the new workbook.
 * @param {Record<WorkbookWorksheetName, Iterable<Partial<Cell>[]> | AsyncIterable<Partial<Cell>[]>>} sheets Object where each key is a sheet name (WorkbookWorksheetName) and the value is an iterable or async iterable of row arrays.
 * @param {CreateOptions} [options] Options for conflict resolution, etc.
 * @returns {Promise<DriveItem & DriveItemRef>} Created DriveItem with reference.
 * @throws {InvalidArgumentError} If the file extension is not supported.
 */
export default async function createWorkbook(parentRef: DriveRef | DriveItemRef, itemPath: DriveItemPath, sheets: Record<WorkbookWorksheetName, Iterable<Partial<Cell>[]> | AsyncIterable<Partial<Cell>[]>>, options: CreateOptions = {}): Promise<DriveItem & DriveItemRef> {
	const extension = extname(itemPath);
	if (extension !== ".xlsx") {
		throw new InvalidArgumentError(`Unsupported file extension: ${extension}. Only .xlsx files are supported for workbook creation.`);
	}

	const {
		conflictBehavior = "fail",
		maxChunkSize = 60 * 1024 * 1024, // 60MB is the largest supported size, minimizing inter-chunk overhead at the expense of large retry blocks
		progress = () => {},
	} = options;

	const localFilePath = pathJoin(tmpdir(), `${randomUUID()}${extension}`);

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
