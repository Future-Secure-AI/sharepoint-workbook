/**
 * Write workbook to Microsoft Sharepoint to a specific path.
 * @module writeWorkbookByPath
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import type { DriveRef } from "microsoft-graph/dist/cjs/models/Drive";
import type { DriveItemPath, DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";
import createDriveItemContent from "microsoft-graph/dist/cjs/operations/driveItem/createDriveItemContent";
import { createReadStream, promises as fs } from "node:fs";
import { extname } from "node:path";
import type { Handle } from "../models/Handle.ts";
import type { WriteOptions } from "../models/Options.ts";
import { streamHighWaterMark } from "../services/streamParameters.ts";
import { getLatestRevisionFilePath } from "../services/workingFolder.ts";

/**
 * Writes a workbook file to Microsoft SharePoint at a given location.
 * @param {Handle} handle Reference to the locally opened workbook.
 * @param {DriveRef | DriveItemRef} parentRef Reference to the parent Drive or DriveItem where the file will be written.
 * @param {DriveItemPath} path Path where the workbook will be written in SharePoint.
 * @param {WriteOptions} [options] Options for writing, such as progress callback.
 * @returns {Promise<void>} Resolves when the workbook has been written.
 */
export default async function writeWorkbookByPath(handle: Handle, parentRef: DriveRef | DriveItemRef, path: DriveItemPath, options: WriteOptions = {}): Promise<DriveItem & DriveItemRef> {
	const extension = extname(path).toLowerCase();
	if (extension !== ".xlsx") {
		throw new Error(`Unsupported file extension: "${extension}". Only .xlsx supported.`);
	}

	const { ifExists = "fail", maxChunkSize, progress } = options;
	const { id } = handle;
	const localPath = await getLatestRevisionFilePath(id);

	const { size } = await fs.stat(localPath);
	const stream = createReadStream(localPath, { highWaterMark: streamHighWaterMark });
	const item = await createDriveItemContent(parentRef, path, stream, size, {
		conflictBehavior: ifExists,
		maxChunkSize,
		progress,
	});
	return item;
}
