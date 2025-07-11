/**
 * Write a workbook.
 * @module writeWorkbook
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
import { getEnvironmentVariable } from "microsoft-graph/dist/cjs/services/environmentVariable";
import { randomUUID } from "node:crypto";
import { createReadStream, createWriteStream, promises as fs } from "node:fs";
import { mkdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { extname, join } from "node:path";
import yauzl, { type Entry, type ZipFile as YauzlZipFile } from "yauzl";
import yazl, { type ZipFile as YazlZipFile } from "yazl";
import { appendRow } from "../services/excelJs.ts";

/**
 * Options for writing a workbook file.
 * @property {WorkbookWorksheetName} [sheetName] Name of the worksheet to create.
 * @property {"fail" | "replace" | "rename"} [ifAlreadyExists] What to do if the file already exists.
 * @property {number} [maxChunkSize] Maximum chunk size for upload (in bytes).
 * @property {(preparedCount: number, writtenCount: number, preparedPerSecond: number, writtenPerSecond: number) =&lt; void} [progress] Progress callback.
 * @property {string} [workingFolder] Working folder for temporary file storage. Defaults to the `WORKING_FOLDER` env, then the OS temporary folder if not set.
 * @property {number} [compressionLevel] Compression level for the output .xlsx zip file (0-9, default 6, )
 */
export type WriteOptions = {
	ifAlreadyExists?: "fail" | "replace" | "rename";
	maxChunkSize?: number;
	progress?: (preparedCount: number, writtenCount: number, preparedPerSecond: number, writtenPerSecond: number) => void;
	workingFolder?: string;
	compressionLevel?: number;
};

/**
 * Writes a workbook (.xlsx) in the specified parent location with the provided rows for multiple sheets.
 * @param {DriveRef | DriveItemRef} parentRef Reference to the parent drive or item where the file will be written.
 * @param {DriveItemPath} itemPath Path (including filename and extension) for the new workbook.
 * @param {Record<WorkbookWorksheetName, Iterable<Partial<Cell>[]> | AsyncIterable<Partial<Cell>[]>>} sheets Object where each key is a sheet name (WorkbookWorksheetName) and the value is an iterable or async iterable of row arrays.
 * @param {WriteOptions} [options] Options for conflict resolution, etc.
 * @returns {Promise<DriveItem & DriveItemRef>} Written DriveItem with reference.
 * @throws {InvalidArgumentError} If the file extension is not supported.
 * @experimental
 */
export default async function writeWorkbook(parentRef: DriveRef | DriveItemRef, itemPath: DriveItemPath, sheets: Record<WorkbookWorksheetName, Iterable<Partial<Cell>[]> | AsyncIterable<Partial<Cell>[]>>, options: WriteOptions = {}): Promise<DriveItem & DriveItemRef> {
	if (extname(itemPath) !== ".xlsx") {
		throw new InvalidArgumentError(`Unsupported file extension: ${extname(itemPath)}. Only .xlsx files are supported for workbook creation.`);
	}

	const { ifAlreadyExists: conflictBehavior = "fail", maxChunkSize = 60 * 1024 * 1024, progress = () => {}, workingFolder = getEnvironmentVariable("WORKING_FOLDER", tmpdir()) as string, compressionLevel = 6 } = options;

	let lastTime = 0;
	let lastPreparedCells = 0;
	let lastWrittenCells = 0;
	let lastPreparedPerSecond = 0;
	let lastWrittenPerSecond = 0;

	const notifyProgress = (force: boolean, preparedCells: number | undefined, writtenCells: number | undefined): void => {
		const time = Date.now();
		const timeDiff = time - lastTime;
		if (force || timeDiff > 1000) {
			if (preparedCells !== undefined) {
				lastPreparedPerSecond = timeDiff ? Math.ceil((preparedCells - lastPreparedCells) / (timeDiff / 1000)) : 0;
				lastPreparedCells = preparedCells;
			}
			if (writtenCells !== undefined) {
				lastWrittenPerSecond = timeDiff ? Math.ceil((writtenCells - lastWrittenCells) / (timeDiff / 1000)) : 0;
				lastWrittenCells = writtenCells;
			}
			lastTime = time;

			progress(lastPreparedCells, lastWrittenCells, lastPreparedPerSecond, lastWrittenPerSecond);
		}
	};

	const scratchFolder = await createScratchFolder(workingFolder);
	const { localWorkbookPath, preparedCells } = await createLocalWorkbook(scratchFolder, sheets, notifyProgress);
	const compressedLocalWorkbookPath = await recompressWorkbook(localWorkbookPath, scratchFolder, compressionLevel);
	return await uploadWorkbook(compressedLocalWorkbookPath, parentRef, itemPath, conflictBehavior, maxChunkSize, preparedCells, notifyProgress);
}

async function createScratchFolder(workingFolder: string): Promise<string> {
	const path = join(workingFolder, randomUUID());
	await mkdir(path, { recursive: true });
	return path;
}

async function createLocalWorkbook(scratchFolder: string, sheets: Record<WorkbookWorksheetName, Iterable<Partial<Cell>[]> | AsyncIterable<Partial<Cell>[]>>, notifyProgress: (force: boolean, preparedCells: number | undefined, writtenCells: number | undefined) => void) {
	const rawFilePath = join(scratchFolder, `raw.xlsx`);
	const rawStream = createWriteStream(rawFilePath);
	const xls = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: rawStream });
	let preparedCells = 0;
	for (const [sheetName, sheetRows] of Object.entries(sheets)) {
		const worksheet = xls.addWorksheet(sheetName);
		for await (const row of sheetRows) {
			appendRow(worksheet, row);
			preparedCells += row.length;
			notifyProgress(false, preparedCells, undefined);
		}
		worksheet.commit();
	}
	await xls.commit();

	notifyProgress(true, undefined, undefined);
	return { localWorkbookPath: rawFilePath, preparedCells };
}

async function recompressWorkbook(inputFilePath: string, scratchFolder: string, compressionLevel: number): Promise<string> {
	if (compressionLevel === 0) {
		return inputFilePath;
	}

	const outputFilePath = join(scratchFolder, `compressed.xlsx`);
	await new Promise<void>((resolve, reject) => {
		yauzl.open(inputFilePath, { lazyEntries: true, autoClose: true }, (err: Error | null, zipfile: YauzlZipFile | undefined) => {
			if (err || !zipfile) return reject(err);
			const zipWriter: YazlZipFile = new yazl.ZipFile();
			const stream = createWriteStream(outputFilePath);
			zipWriter.outputStream.pipe(stream).on("close", resolve).on("error", reject);
			zipfile.readEntry();
			zipfile.on("entry", (entry: Entry) => {
				zipfile.openReadStream(entry, (err: Error | null, readStream: NodeJS.ReadableStream | undefined) => {
					if (err || !readStream) return reject(err);
					zipWriter.addReadStream(readStream, entry.fileName, {
						compress: true,
						compressionLevel,
						mtime: entry.getLastModDate ? entry.getLastModDate() : new Date(),
						mode: entry.externalFileAttributes >>> 16,
					});
					zipfile.readEntry();
				});
			});
			zipfile.on("end", () => {
				zipWriter.end();
			});
			zipfile.on("error", reject);
		});
	});
	return outputFilePath;
}

async function uploadWorkbook(
	compressedFilePath: string,
	parentRef: DriveRef | DriveItemRef,
	itemPath: DriveItemPath,
	conflictBehavior: "fail" | "replace" | "rename",
	maxChunkSize: number,
	preparedCells: number,
	notifyProgress: (force: boolean, preparedCells: number | undefined, writtenCells: number | undefined) => void,
): Promise<DriveItem & DriveItemRef> {
	const { size } = await fs.stat(compressedFilePath);
	const stream = createReadStream(compressedFilePath, { highWaterMark: 1024 * 1024 });
	const item = await createDriveItemContent(parentRef, itemPath, stream, size, {
		conflictBehavior,
		maxChunkSize,
		progress: (bytes) => {
			const writtenCells = Math.ceil((bytes / size) * preparedCells);
			notifyProgress(false, undefined, writtenCells);
		},
	});
	notifyProgress(true, undefined, undefined);
	return item;
}
