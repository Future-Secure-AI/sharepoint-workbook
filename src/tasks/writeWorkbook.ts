/**
 * Write a workbook.
 * @module writeWorkbook
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import ExcelJS from "exceljs";
import InvalidArgumentError from "microsoft-graph/dist/cjs/errors/InvalidArgumentError";
import type { DriveRef } from "microsoft-graph/dist/cjs/models/Drive";
import type { DriveItemPath, DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";
import createDriveItemContent from "microsoft-graph/dist/cjs/operations/driveItem/createDriveItemContent";
import { getEnvironmentVariable } from "microsoft-graph/dist/cjs/services/environmentVariable";
import { randomUUID } from "node:crypto";
import { createReadStream, createWriteStream, promises as fs } from "node:fs";
import { mkdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { extname, join } from "node:path";
import yauzl, { type Entry, type ZipFile as YauzlZipFile } from "yauzl";
import yazl, { type ZipFile as YazlZipFile } from "yazl";
import type { WriteWorksheet } from "../models/Worksheet.ts";
import { appendRow } from "../services/excelJs.ts";

/**
 * Progress information for workbook writing operations.
 * @typedef {Object} WriteProgress
 * @property {number} prepared Number of cells prepared for writing
 * @property {number} written Number of cells written to the destination
 * @property {number} compressionRatio Ratio of compressed file size to original file size (0 to 1, where 1 is no compression)
 * @property {number} preparedPerSecond Number of cells prepared per second
 * @property {number} writtenPerSecond Number of cells written per second
 */
export type WriteProgress = {
	prepared: number;
	written: number;
	compressionRatio: number;
	preparedPerSecond: number;
	writtenPerSecond: number;
};

/**
 * Options for writing a workbook file.
 * @typedef {Object} WriteOptions
 * @property {"fail" | "replace" | "rename"} [ifAlreadyExists] What to do if the file already exists.
 * @property {number} [maxChunkSize] Maximum chunk size for upload (in bytes).
 * @property {(update: WriteProgress) => void} [progress] Progress callback.
 * @property {string} [workingFolder] Working folder for temporary file storage. Defaults to the `WORKING_FOLDER` env, then the OS temporary folder if not set.
 * @property {number} [compressionLevel] Compression level for the output .xlsx zip file (0-9, default 6)
 */
export type WriteOptions = {
	ifAlreadyExists?: "fail" | "replace" | "rename";
	maxChunkSize?: number;
	progress?: (update: WriteProgress) => void;
	workingFolder?: string;
	compressionLevel?: number;
};

/**
 * Writes a workbook (.xlsx) in the specified parent location with the provided rows for multiple sheets.
 * @param {DriveRef | DriveItemRef} parentRef Reference to the parent drive or item where the file will be written.
 * @param {DriveItemPath} itemPath Path (including filename and extension) for the new workbook.
 * @param {AsyncIterator<WriteWorksheet>} worksheets Worksheets to be written.
 * @param {WriteOptions} [options] Options for conflict resolution, etc.
 * @returns {Promise<DriveItem & DriveItemRef>} Written DriveItem with reference.
 * @throws {InvalidArgumentError} If the file extension is not supported.
 * @experimental
 */
export default async function writeWorkbook(parentRef: DriveRef | DriveItemRef, itemPath: DriveItemPath, worksheets: Iterable<WriteWorksheet> | AsyncIterable<WriteWorksheet>, options: WriteOptions = {}): Promise<DriveItem & DriveItemRef> {
	if (extname(itemPath) !== ".xlsx") {
		throw new InvalidArgumentError(`Unsupported file extension: ${extname(itemPath)}. Only .xlsx files are supported for workbook creation.`);
	}

	const { ifAlreadyExists: conflictBehavior = "fail", maxChunkSize = 60 * 1024 * 1024, progress = () => {}, workingFolder = getEnvironmentVariable("WORKING_FOLDER", tmpdir()) as string, compressionLevel = 6 } = options;

	const state = {
		time: Date.now(),
		prepared: 0,
		written: 0,
		compressionRatio: 0,
		preparedPerSecond: 0,
		writtenPerSecond: 0,
	};

	const reportProgress = (prepared?: number, compressionRatio?: number, written?: number, force = false): void => {
		const now = Date.now();
		const elapsed = now - state.time;

		state.compressionRatio = compressionRatio ?? state.compressionRatio;

		if (force || elapsed > 1000) {
			if (prepared !== undefined) {
				state.preparedPerSecond = elapsed ? Math.ceil((prepared - state.prepared) / (elapsed / 1000)) : 0;
				state.prepared = prepared ?? state.prepared;
			}
			if (written !== undefined) {
				state.writtenPerSecond = elapsed ? Math.ceil((written - state.written) / (elapsed / 1000)) : 0;
				state.written = written ?? state.written;
			}

			state.time = now;
			progress(state);
		}
	};

	const scratchFolder = await createScratchFolder(workingFolder);
	const { localWorkbookPath, preparedCells } = await createLocalWorkbook(scratchFolder, worksheets, reportProgress);
	const compressedLocalWorkbookPath = await recompressWorkbook(localWorkbookPath, scratchFolder, compressionLevel, reportProgress);
	return await uploadWorkbook(compressedLocalWorkbookPath, parentRef, itemPath, conflictBehavior, maxChunkSize, preparedCells, reportProgress);
}

async function createScratchFolder(workingFolder: string): Promise<string> {
	const path = join(workingFolder, randomUUID());
	await mkdir(path, { recursive: true });
	return path;
}

async function createLocalWorkbook(scratchFolder: string, worksheets: Iterable<WriteWorksheet> | AsyncIterable<WriteWorksheet>, notifyProgress: (prepared: number | undefined, compressionRation: number | undefined, written: number | undefined, force: boolean) => void) {
	const rawFilePath = join(scratchFolder, `raw.xlsx`);
	const rawStream = createWriteStream(rawFilePath);
	const xls = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: rawStream });
	let preparedCells = 0;
	for await (const { name, rows } of worksheets) {
		const worksheet = xls.addWorksheet(name);
		for await (const row of rows) {
			appendRow(worksheet, row);
			preparedCells += row.length;
			notifyProgress(preparedCells, undefined, undefined, false);
		}
		worksheet.commit();
	}
	await xls.commit();

	notifyProgress(undefined, undefined, undefined, true);
	return { localWorkbookPath: rawFilePath, preparedCells };
}

async function recompressWorkbook(inputFilePath: string, scratchFolder: string, compressionLevel: number, notifyProgress: (prepared: number | undefined, compressionRation: number | undefined, written: number | undefined, force: boolean) => void): Promise<string> {
	if (compressionLevel === 0) {
		notifyProgress(undefined, 1, undefined, true);
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

	const { size: inputSize } = await fs.stat(inputFilePath);
	const { size: outputSize } = await fs.stat(outputFilePath);
	notifyProgress(undefined, outputSize / inputSize, undefined, true);
	return outputFilePath;
}

async function uploadWorkbook(
	compressedFilePath: string,
	parentRef: DriveRef | DriveItemRef,
	itemPath: DriveItemPath,
	conflictBehavior: "fail" | "replace" | "rename",
	maxChunkSize: number,
	preparedCells: number,
	notifyProgress: (prepared: number | undefined, compressionRation: number | undefined, written: number | undefined, force: boolean) => void,
): Promise<DriveItem & DriveItemRef> {
	const { size } = await fs.stat(compressedFilePath);
	const stream = createReadStream(compressedFilePath, { highWaterMark: 1024 * 1024 });
	const item = await createDriveItemContent(parentRef, itemPath, stream, size, {
		conflictBehavior,
		maxChunkSize,
		progress: (bytes) => {
			const writtenCells = Math.ceil((bytes / size) * preparedCells);
			notifyProgress(undefined, undefined, writtenCells, false);
		},
	});
	notifyProgress(undefined, undefined, undefined, true);
	return item;
}
