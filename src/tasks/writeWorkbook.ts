import createDriveItemContent from "microsoft-graph/createDriveItemContent";
import type { DriveItemPath } from "microsoft-graph/DriveItem";
import getDriveItem from "microsoft-graph/getDriveItem";
import { createReadStream } from "node:fs";
import { stat } from "node:fs/promises";
import type { OpenRef } from "../models/Open.ts";
import type { WriteOptions } from "../models/WriteOptions.ts";
import { highWaterMark, maxChunkSize } from "../services/streamParameters.ts";
import { getLatest } from "../services/workingFolder.ts";

export default async function writeWorkbook(openRef: OpenRef, options: WriteOptions = {}): Promise<void> {
	const { itemRef } = openRef;
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

	const { ifExists = "fail", progress } = options;
	const { id } = openRef;
	const localPath = await getLatest(id);

	const { size } = await stat(localPath);
	const stream = createReadStream(localPath, { highWaterMark });
	await createDriveItemContent(parentRef, path, stream, size, {
		conflictBehavior: ifExists,
		maxChunkSize,
		progress,
	});
}
