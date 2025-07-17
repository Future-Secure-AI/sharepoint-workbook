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
import type { WriteOptions } from "../models/Options.ts";
import type { Workbook } from "../models/Workbook.ts";
import { streamHighWaterMark } from "../services/streamParameters.ts";
import { getTemporaryFilePath } from "../services/temporaryFile.ts";

/**
 * Writes a workbook file to Microsoft SharePoint at a given location.
 * @param {Workbook} workbook Reference to the locally opened workbook.
 * @param {DriveRef | DriveItemRef} parentRef Reference to the parent Drive or DriveItem where the file will be written.
 * @param {DriveItemPath} path Path where the workbook will be written in SharePoint.
 * @param {WriteOptions} [options] Options for writing, such as progress callback.
 * @returns {Promise<void>} Resolves when the workbook has been written.
 */
export default async function saveWorkbookAs(workbook: Workbook, parentRef: DriveRef | DriveItemRef, path: DriveItemPath, options: WriteOptions = {}): Promise<DriveItem & DriveItemRef> {
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
