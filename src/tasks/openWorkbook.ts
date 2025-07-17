/**
 * Open workbook from Microsoft SharePoint.
 * @module openWorkbook
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import AsposeCells from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/dist/cjs/errors/InvalidArgumentError";
import type { DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";
import getDriveItem from "microsoft-graph/dist/cjs/operations/driveItem/getDriveItem";
import streamDriveItemContent from "microsoft-graph/dist/cjs/operations/driveItem/streamDriveItemContent";
import { createWriteStream } from "node:fs";
import { unlink } from "node:fs/promises";
import { extname } from "node:path";
import { pipeline } from "node:stream/promises";
import type { Handle } from "../models/Handle.ts";
import type { LocalFilePath } from "../models/LocalFilePath.ts";
import type { ReadOptions } from "../models/Options.ts";
import { getTemporaryFilePath } from "../services/temporaryFile.ts";

/**
 * Reads a workbook file (.xlsx or .csv) from a Microsoft Graph.
 * @param {DriveItemRef & Partial<DriveItem>} remoteItemRef - Reference to the DriveItem to read from.
 * @param {ReadOptions} [options] - Options for reading, such as default worksheet name for CSV.
 * @returns {Promise<Handle>} Reference to the locally opened workbook.
 * @throws {InvalidArgumentError} If the file extension is not supported.
 */
export default async function openWorkbook(remoteItemRef: DriveItemRef | (DriveItemRef & DriveItem), options: ReadOptions = {}): Promise<Handle> {
	const { progress = () => {} } = options;

	const extension = await getDriveItemFileExtension(remoteItemRef);
	const localFilePath = await getTemporaryFilePath(extension);

	await downloadFile(remoteItemRef, localFilePath, progress);
	const workbook = openFile(localFilePath, extension);
	await unlink(localFilePath);

	if (extension === ".xlsx") {
		return {
			workbook,
			remoteItemRef, // Only xlsx files can be logically overwritten, other formats will have the incorrect file extension
		};
	} else {
		return {
			workbook,
		};
	}
}

async function getDriveItemFileExtension(remoteItemRef: DriveItemRef | (DriveItemRef & DriveItem)) {
	let name: string | undefined;
	if ("name" in remoteItemRef && typeof remoteItemRef.name === "string") {
		name = remoteItemRef.name;
	} else {
		const driveItem = await getDriveItem(remoteItemRef);
		name = driveItem?.name || "";
	}

	if (!name) throw new InvalidArgumentError("DriveItem does not have a valid 'name' property to determine file extension.");

	const extension = extname(name).toLowerCase();
	if (!extension) throw new InvalidArgumentError(`Could not determine file extension from name: '${name}'`);

	return extension;
}

function openFile(localFilePath: LocalFilePath, extension: string): AsposeCells.Workbook {
	if (extension === ".csv") {
		return new AsposeCells.Workbook(localFilePath, new AsposeCells.TxtLoadOptions(AsposeCells.LoadFormat.Csv));
	} else if (extension === ".tsv") {
		return new AsposeCells.Workbook(localFilePath, new AsposeCells.TxtLoadOptions(AsposeCells.LoadFormat.Tsv));
		// TODO: Probably more text files that need manual handling
	} else {
		return new AsposeCells.Workbook(localFilePath);
	}
}

async function downloadFile(itemRef: DriveItemRef, localFilePath: LocalFilePath, progress?: (bytes: number) => void): Promise<void> {
	const sourceStream = await streamDriveItemContent(itemRef);
	const destinationStream = createWriteStream(localFilePath);
	let bytesProcessed = 0;
	let lastProgressTime = 0;
	if (progress) {
		sourceStream.on("data", (chunk) => {
			bytesProcessed += chunk.length;
			const now = Date.now();
			if (now - lastProgressTime >= 1000) {
				progress(bytesProcessed);
				lastProgressTime = now;
			}
		});
	}
	await pipeline(sourceStream, destinationStream);
}
