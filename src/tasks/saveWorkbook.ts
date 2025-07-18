/**
 * Write opened workbook back to Microsoft SharePoint.
 * @module saveWorkbook
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import type { DriveItemPath, DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";
import { getDriveItemParent } from "microsoft-graph/driveItem";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import MissingPathError from "../errors/MissingPathError.ts";
import type { Workbook } from "../models/Workbook.ts";
import saveWorkbookAs, { type SaveWorkbookOptions } from "./saveWorkbookAs.ts";

/**
 * Write a locally opened workbook back to Microsoft SharePoint, overwriting the previous file.
 * @param {Workbook} handle Reference to the locally opened workbook, must include an itemRef for overwrite.
 * @param {SaveWorkbookOptions} [options] Options for writing, such as conflict behavior, chunk size, and progress callback.
 * @returns {Promise<void>} Resolves when the upload is complete.
 * @throws {Error} If the workbook cannot be overwritten or required metadata is missing.
 */
export default async function saveWorkbook(handle: Workbook, options: SaveWorkbookOptions = {}): Promise<DriveItemRef & DriveItem> {
	const { ifExists = "replace", maxChunkSize, progress } = options;

	const remoteItem = handle.remoteItem;
	if (!remoteItem) throw new MissingPathError("Workbook hasn't been 'saved as', so no path is known to overwrite. Use `saveWorkbookAs` instead.");

	const parentRef = getDriveItemParent(remoteItem);
	const path = getDriveItemPathWithinParent(remoteItem);

	return await saveWorkbookAs(handle, parentRef, path, { ifExists, maxChunkSize, progress });
}

function getDriveItemPathWithinParent(remoteItem: DriveItem): DriveItemPath {
	const name = remoteItem.name;
	if (!name) throw new InvalidArgumentError("Missing drive item name.");
	const path = `/${name}` as DriveItemPath;
	return path;
}
