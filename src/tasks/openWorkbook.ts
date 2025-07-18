/**
 * Reading a workbook from SharePoint by path.
 * @module openWorkbookByPath
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import AsposeCells from "aspose.cells.node";
import NotFoundError from "microsoft-graph/dist/cjs/errors/NotFoundError";
import type { DriveRef } from "microsoft-graph/dist/cjs/models/Drive";
import type { DriveItemName, DriveItemPath, DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";
import getDriveItemByPath from "microsoft-graph/dist/cjs/operations/driveItem/getDriveItemByPath";
import streamDriveItemContent from "microsoft-graph/dist/cjs/operations/driveItem/streamDriveItemContent";
import iterateDriveItems from "microsoft-graph/dist/cjs/tasks/iterateDriveItems";
import { createWriteStream } from "node:fs";
import { unlink } from "node:fs/promises";
import { extname } from "node:path";
import { pipeline } from "node:stream/promises";
import picomatch from "picomatch";
import type { LocalFilePath } from "../models/LocalFilePath.ts";
import type { Workbook } from "../models/Workbook.ts";
import { getTemporaryFilePath } from "../services/temporaryFile.ts";

/**
 * Options for reading a workbook file.
 * @property {WorkbookWorksheetName} [defaultWorksheetName] Default worksheet name to use when importing a CSV file.
 * @property {(bytes: number): void} [progress] Progress callback, receives bytes processed.
 */
export type OpenWorkbookOptions = {
	progress?: (bytes: number) => void;
};

/**
 * Reads a workbook file from a SharePoint drive by its path, supporting wildcards in the filename.
 * @param {DriveRef | DriveItemRef} parentRef - Reference to the parent drive or folder.
 * @param {DriveItemPath} itemPath - Path to the file, may include wildcards in the filename.
 * @returns {Promise<Workbook>} Reference to the locally opened workbook.
 * @throws {Error} If the file path is invalid or no matching file is found.
 * @remarks It could be any file that is:
 * - No more than 250GB
 * - No more than 4x amount of available server memory
 * - No more than the configured Node memory limit (default 4GB) less what's already used
 * - Is a supported file type https://docs.aspose.com/cells/cpp/supported-file-formats/
 */
export default async function openWorkbook(parentRef: DriveRef | DriveItemRef, itemPath: DriveItemPath, options: OpenWorkbookOptions = {}): Promise<Workbook> {
	const { progress = () => {} } = options;

	const { folderPath, fileName: filePattern } = decomposePath(itemPath);
	const folder = await getDriveItemByPath(parentRef, folderPath);
	const items = iterateDriveItems(folder);
	const remoteItem = await matchFile(filePattern, items);

	const name = remoteItem.name ?? "";
	const extension = extname(name).toLowerCase();

	const tempFile = await getTemporaryFilePath(extension);
	await downloadFile(remoteItem, tempFile, progress);
	const workbook = openFile(tempFile, extension) as Workbook;
	await unlink(tempFile);

	workbook.remoteItem = remoteItem;
	return workbook as Workbook;
}

function decomposePath(itemPath: DriveItemPath): { folderPath: DriveItemPath; fileName: DriveItemName } {
	if (itemPath === "/") {
		throw new Error(`Invalid file path: "${itemPath}". It must not be just a forward slash ("/").`);
	}
	const pos = itemPath.lastIndexOf("/");
	if (pos === -1) {
		throw new Error(`Invalid file path: "${itemPath}". It must contain at least one forward slash ("/").`);
	}
	if (pos === itemPath.length - 1) {
		throw new Error(`Invalid file path: "${itemPath}". It must not end with a forward slash ("/").`);
	}

	const folderPath = itemPath.slice(0, pos + 1) as DriveItemPath;
	const fileName = itemPath.slice(pos + 1) as DriveItemName;

	return {
		folderPath,
		fileName,
	};
}

async function matchFile(filePattern: DriveItemName, items: AsyncIterable<DriveItem & DriveItemRef>): Promise<DriveItem & DriveItemRef> {
	const isMatch = picomatch(filePattern, { nocase: true });

	for await (const item of items) {
		const name = item.name ?? "";
		if (isMatch(name)) {
			return item;
		}
	}

	throw new NotFoundError(`No file matching pattern "${filePattern}" found in the specified folder.`);
}

function openFile(localFilePath: LocalFilePath, extension: string): AsposeCells.Workbook {
	if (extension === ".csv") {
		return new AsposeCells.Workbook(localFilePath, new AsposeCells.TxtLoadOptions(AsposeCells.LoadFormat.Csv));
	} else if (extension === ".tsv") {
		return new AsposeCells.Workbook(localFilePath, new AsposeCells.TxtLoadOptions(AsposeCells.LoadFormat.Tsv));
		// TODO: Probably more text files that need manual handling
	} else {
		return new AsposeCells.Workbook(localFilePath);
	}
}

async function downloadFile(itemRef: DriveItemRef, localFilePath: LocalFilePath, progress?: (bytes: number) => void): Promise<void> {
	const sourceStream = await streamDriveItemContent(itemRef);
	const destinationStream = createWriteStream(localFilePath);
	let bytesProcessed = 0;
	let lastProgressTime = 0;
	if (progress) {
		sourceStream.on("data", (chunk) => {
			bytesProcessed += chunk.length;
			const now = Date.now();
			if (now - lastProgressTime >= 1000) {
				progress(bytesProcessed);
				lastProgressTime = now;
			}
		});
	}
	await pipeline(sourceStream, destinationStream);
}
