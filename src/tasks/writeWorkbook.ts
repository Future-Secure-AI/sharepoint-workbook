/**
 * Write a locally opened workbook back to Microsoft SharePoint.
 * @module writeWorkbook
 * @category Tasks
 */

import type { DriveItemPath } from "microsoft-graph/dist/cjs/models/DriveItem";
import createDriveItemContent from "microsoft-graph/dist/cjs/operations/driveItem/createDriveItemContent";
import getDriveItem from "microsoft-graph/dist/cjs/operations/driveItem/getDriveItem";
import { createReadStream } from "node:fs";
import { stat } from "node:fs/promises";
import type { Handle } from "../models/Handle.ts";
import type { WriteOptions } from "../models/WriteOptions.ts";
import { streamHighWaterMark } from "../services/streamParameters.ts";
import { getLatestRevisionFilePath } from "../services/workingFolder.ts";

/**
 * Write a locally opened workbook back to Microsoft SharePoint, overwriting the previous file.
 * @param {Handle} hdl Reference to the locally opened workbook, must include an itemRef for overwrite.
 * @param {WriteOptions} [options] Options for writing, such as conflict behavior, chunk size, and progress callback.
 * @returns {Promise<void>} Resolves when the upload is complete.
 * @throws {Error} If the workbook cannot be overwritten or required metadata is missing.
 */
export default async function writeWorkbook(hdl: Handle, options: WriteOptions = {}): Promise<void> {
	const { itemRef } = hdl;
	if (!itemRef) {
		throw new Error("Workbook not over-writable. Use `writeWorkbookByPath` instead.");
	}

	const item = await getDriveItem(itemRef);

	const parentId = item.parentReference?.id;
	if (!parentId) {
		throw new Error("Parent reference not found for the item.");
	}

	const parentRef = {
		...itemRef,
		id: parentId,
	};

	const name = item.name;
	if (!name) {
		throw new Error("Item name not found.");
	}
	const path = `/${name}` as DriveItemPath;

	const { ifExists = "replace", maxChunkSize, progress } = options;
	const { id } = hdl;
	const localPath = await getLatestRevisionFilePath(id);

	const { size } = await stat(localPath);
	const stream = createReadStream(localPath, { highWaterMark: streamHighWaterMark });
	await createDriveItemContent(parentRef, path, stream, size, {
		conflictBehavior: ifExists,
		maxChunkSize,
		progress,
	});
}
