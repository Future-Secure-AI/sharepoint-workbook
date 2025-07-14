/**
 * Read workbook from Microsoft SharePoint.
 * @module readWorkbook
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import ExcelJS from "exceljs";
import { parse } from "fast-csv";
import InvalidArgumentError from "microsoft-graph/dist/cjs/errors/InvalidArgumentError";
import type { DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";
import getDriveItem from "microsoft-graph/dist/cjs/operations/driveItem/getDriveItem";
import streamDriveItemContent from "microsoft-graph/dist/cjs/operations/driveItem/streamDriveItemContent";
import { defaultWorkbookWorksheetName } from "microsoft-graph/dist/cjs/services/workbookWorksheet";
import { createWriteStream } from "node:fs";
import { extname } from "node:path";
import { pipeline } from "node:stream/promises";
import type { Handle } from "../models/Handle.ts";
import type { ReadOptions } from "../models/ReadOptions.ts";
import { createHandleId, getNextRevisionFilePath } from "../services/workingFolder.ts";

/**
 * Reads a workbook file (.xlsx or .csv) from a Microsoft Graph.
 * @param {DriveItemRef & Partial<DriveItem>} itemRef - Reference to the DriveItem to read from.
 * @param {ReadOptions} [options] - Options for reading, such as default worksheet name for CSV.
 * @returns {Promise<Handle>} Reference to the locally opened workbook.
 * @throws {InvalidArgumentError} If the file extension is not supported.
 */
export default async function readWorkbook(itemRef: DriveItemRef & Partial<DriveItem>, options: ReadOptions = {}): Promise<Handle> {
	const { defaultWorksheetName = defaultWorkbookWorksheetName, progress = () => {} } = options;

	const id = createHandleId();
	let name = itemRef.name;
	if (!name) name = (await getDriveItem(itemRef)).name ?? "";

	const extension = extname(name).toLowerCase();
	const targetFileName = await getNextRevisionFilePath(id);
	const stream = await streamDriveItemContent(itemRef);

	let bytesProcessed = 0;
	let lastProgressTime = 0;
	stream.on("data", (chunk) => {
		bytesProcessed += chunk.length;
		const now = Date.now();
		if (now - lastProgressTime >= 1000) {
			progress(bytesProcessed);
			lastProgressTime = now;
		}
	});

	if (extension === ".xlsx") {
		await pipeline(stream, createWriteStream(targetFileName));
		return { id, itemRef };
	}

	if (extension === ".csv") {
		const xlsx = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: targetFileName });
		const worksheet = xlsx.addWorksheet(defaultWorksheetName);
		await new Promise<void>((resolve, reject) => {
			stream
				.pipe(parse({ headers: false }))
				.on("error", reject)
				.on("data", (row: unknown[]) => worksheet.addRow(row).commit())
				.on("end", async () => {
					worksheet.commit();
					await xlsx.commit();
					resolve();
				});
		});
		return { id };
	}

	throw new InvalidArgumentError(`Unsupported file extension "${extension}".`);
}
