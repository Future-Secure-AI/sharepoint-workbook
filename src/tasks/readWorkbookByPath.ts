/**
 * Reading a workbook from SharePoint by path.
 * @module readWorkbookByPath
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import NotFoundError from "microsoft-graph/dist/cjs/errors/NotFoundError";
import type { DriveRef } from "microsoft-graph/dist/cjs/models/Drive";
import type { DriveItemName, DriveItemPath, DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";
import getDriveItemByPath from "microsoft-graph/dist/cjs/operations/driveItem/getDriveItemByPath";
import iterateDriveItems from "microsoft-graph/dist/cjs/tasks/iterateDriveItems";
import picomatch from "picomatch";
import type { Handle } from "../models/Handle.ts";
import type { ReadOptions } from "../models/Options.ts";
import readWorkbook from "./readWorkbook.ts";

/**
 * Reads a workbook file from a SharePoint drive by its path, supporting wildcards in the filename.
 * @param {DriveRef | DriveItemRef} parentRef - Reference to the parent drive or folder.
 * @param {DriveItemPath} itemPath - Path to the file, may include wildcards in the filename.
 * @returns {Promise<Handle>} Reference to the locally opened workbook.
 * @throws {Error} If the file path is invalid or no matching file is found.
 */
export default async function readWorkbookByPath(parentRef: DriveRef | DriveItemRef, itemPath: DriveItemPath, options: ReadOptions = {}): Promise<Handle> {
	const { folderPath, fileName: filePattern } = decomposePath(itemPath);
	const folder = await getDriveItemByPath(parentRef, folderPath);
	const items = iterateDriveItems(folder);
	const item = await matchFile(filePattern, items);
	
	return await readWorkbook(item, options);
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
