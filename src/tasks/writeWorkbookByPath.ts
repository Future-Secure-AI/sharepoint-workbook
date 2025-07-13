/**
 * Write workbook to Microsoft Sharepoint to a specific path.
 * @module writeWorkbookByPath
 * @category Tasks
 */

import type { DriveRef } from "microsoft-graph/dist/cjs/models/Drive";
import type { DriveItemPath, DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";
import createDriveItemContent from "microsoft-graph/dist/cjs/operations/driveItem/createDriveItemContent";
import { createReadStream, promises as fs } from "node:fs";
import type { OpenRef } from "../models/Open.ts";
import { highWaterMark, maxChunkSize } from "../services/streamParameters.ts";
import { getLatest } from "../services/workingFolder.ts";
import type { WriteOptions } from "../models/WriteOptions.ts";

/**
 * Writes a workbook file to Microsoft SharePoint.=
 * @param {OpenRef} openRef Reference to the locally opened workbook.
 * @param {DriveRef | DriveItemRef} parentRef Reference to the parent Drive or DriveItem where the file will be written.
 * @param {DriveItemPath} path Path where the workbook will be written in SharePoint.
 * @param {WriteOptions} [options] Options for writing, such as progress callback.
 * @returns {Promise<void>} Resolves when the workbook has been written.
 */
export default async function writeWorkbookByPath(openRef: OpenRef, parentRef: DriveRef | DriveItemRef, path: DriveItemPath, options: WriteOptions = {}): Promise<void> {
	const { ifExists = "fail", progress } = options;
	const { id } = openRef;
	const localPath = await getLatest(id);

	const { size } = await fs.stat(localPath);
	const stream = createReadStream(localPath, { highWaterMark });
	await createDriveItemContent(parentRef, path, stream, size, {
		conflictBehavior: ifExists,
		maxChunkSize,
		progress,
	});
}
