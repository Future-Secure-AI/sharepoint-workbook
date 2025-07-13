import type { DriveItem } from "@microsoft/microsoft-graph-types";
import type { DriveRef } from "microsoft-graph/Drive";
import type { DriveItemName, DriveItemPath, DriveItemRef } from "microsoft-graph/DriveItem";
import getDriveItemByPath from "microsoft-graph/getDriveItemByPath";
import iterateDriveItems from "microsoft-graph/iterateDriveItems";
import NotFoundError from "microsoft-graph/NotFoundError";
import picomatch from "picomatch";
import type { OpenRef } from "../models/Open.ts";
import readWorkbook from "./readWorkbook.ts";

export default async function readWorkbookByPath(parentRef: DriveRef | DriveItemRef, itemPath: DriveItemPath): Promise<OpenRef> {
	const { folderPath, fileName: filePattern } = decomposePath(itemPath);
	const folder = await getDriveItemByPath(parentRef, folderPath);
	const items = iterateDriveItems(folder);
	const item = await matchFile(filePattern, items);
	return await readWorkbook(item);
}

function decomposePath(itemPath: DriveItemPath): { folderPath: DriveItemPath; fileName: DriveItemName } {
	const pos = itemPath.lastIndexOf("/");
	if (pos === -1) {
		throw new Error(`Invalid file path: "${itemPath}". It must contain at least one forward slash ("/").`);
	}
	if (pos === 0) {
		throw new Error(`Invalid file path: "${itemPath}". It must not be just a forward slash ("/").`);
	}
	if (pos === itemPath.length - 1) {
		throw new Error(`Invalid file path: "${itemPath}". It must not end with a forward slash ("/").`);
	}

	const folderPath = itemPath.slice(0, pos) as DriveItemPath;
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
