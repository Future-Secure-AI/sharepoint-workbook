/**
 * Write workbook to Microsoft Sharepoint to a specific path.
 * @module saveWorkbookAs
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import AsposeCells from "aspose.cells.node";
import type { DriveRef } from "microsoft-graph/dist/cjs/models/Drive";
import type { DriveItemPath, DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";
import createDriveItemContent from "microsoft-graph/dist/cjs/operations/driveItem/createDriveItemContent";
import { createReadStream, promises as fs } from "node:fs";
import { unlink } from "node:fs/promises";
import { extname } from "node:path";
import type { Workbook } from "../models/Workbook.ts";
import { streamHighWaterMark } from "../services/streamParameters.ts";
import { getTemporaryFilePath } from "../services/temporaryFile.ts";

/**
 * Options for writing a workbook file.
 * @property {"fail" | "replace" | "rename"} [ifExists] Behavior if the file already exists.
 * @property {(bytes: number): void} [progress] Progress callback, receives bytes processed.
 * @property {number} [maxChunkSize] Maximum chunk size in bytes for writing.
 */
export type SaveWorkbookOptions = {
	ifExists?: "fail" | "replace" | "rename";
	progress?: (bytes: number) => void;
	maxChunkSize?: number;
};

/**
 * Writes a workbook file to Microsoft SharePoint at a given location.
 * @param {Workbook} workbook Reference to the locally opened workbook.
 * @param {DriveRef | DriveItemRef} parentRef Reference to the parent Drive or DriveItem where the file will be written.
 * @param {DriveItemPath} path Path where the workbook will be written in SharePoint.
 * @param {SaveWorkbookOptions} [options] Options for writing, such as progress callback.
 * @returns {Promise<void>} Resolves when the workbook has been written.
 * @remarks See https://docs.aspose.com/cells/cpp/supported-file-formats/ for supported file formats. It cannot exceed SharePoint's file size limit of 250GB.
 * For size indication, a particular 700MB CSV file compresses down to about:
 *  - ~100MB XLSX
 *  - ~30MB XLSB
 *  - ~12MB XLS
 */
export default async function saveWorkbookAs(workbook: Workbook, parentRef: DriveRef | DriveItemRef, path: DriveItemPath, options: SaveWorkbookOptions = {}): Promise<DriveItem & DriveItemRef> {
	const extension = extname(path).toLowerCase();

	const { ifExists = "fail", maxChunkSize, progress } = options;

	const tempFilePath = await getTemporaryFilePath(extension);
	workbook.save(tempFilePath, AsposeCells.SaveFormat.Auto);

	const { size } = await fs.stat(tempFilePath);
	const stream = createReadStream(tempFilePath, { highWaterMark: streamHighWaterMark });
	const remoteItem = await createDriveItemContent(parentRef, path, stream, size, {
		conflictBehavior: ifExists,
		maxChunkSize,
		progress,
	});

	workbook.remoteItem = remoteItem;
	await unlink(tempFilePath);
	return remoteItem;
}
