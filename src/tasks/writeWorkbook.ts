/**
 * Write a locally opened workbook back to Microsoft SharePoint.
 * @module writeWorkbook
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import type { DriveItemPath, DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";
import getDriveItem from "microsoft-graph/dist/cjs/operations/driveItem/getDriveItem";
import { getDriveItemParent } from "microsoft-graph/driveItem";
import type { Handle } from "../models/Handle.ts";
import type { WriteOptions } from "../models/Options.ts";
import writeWorkbookByPath from "./writeWorkbookByPath.ts";

/**
 * Write a locally opened workbook back to Microsoft SharePoint, overwriting the previous file.
 * @param {Handle} handle Reference to the locally opened workbook, must include an itemRef for overwrite.
 * @param {WriteOptions} [options] Options for writing, such as conflict behavior, chunk size, and progress callback.
 * @returns {Promise<void>} Resolves when the upload is complete.
 * @throws {Error} If the workbook cannot be overwritten or required metadata is missing.
 */
export default async function writeWorkbook(handle: Handle, options: WriteOptions = {}): Promise<DriveItemRef & DriveItem> {
	const { remoteItemRef: itemRef } = handle;
	if (!itemRef) {
		throw new Error("Workbook not over-writable. Use `writeWorkbookByPath` instead.");
	}

	const item = await getDriveItem(itemRef);
	const parentRef = getDriveItemParent(item);

	const name = item.name;
	if (!name) {
		throw new Error("Item name not found.");
	}
	const path = `/${name}` as DriveItemPath;

	const { ifExists = "replace", maxChunkSize, progress } = options;

	return await writeWorkbookByPath(handle, parentRef, path, { ifExists, maxChunkSize, progress });
}
