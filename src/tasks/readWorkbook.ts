/**
 * Read a workbook from Microsoft SharePoint.
 * @module readWorkbook
 * @category Tasks
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import ExcelJS from "exceljs";
import { parse } from "fast-csv";
import type { DriveItemRef } from "microsoft-graph/DriveItem";
import getDriveItem from "microsoft-graph/getDriveItem";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import streamDriveItemContent from "microsoft-graph/streamDriveItemContent";
import { defaultWorkbookWorksheetName } from "microsoft-graph/workbookWorksheet";
import type { WorkbookWorksheetName } from "microsoft-graph/WorkbookWorksheet";
import { createWriteStream } from "node:fs";
import { extname, join } from "node:path";
import { pipeline } from "node:stream/promises";
import type { OpenRef } from "../models/Open.ts";
import { createOpenId, getWorkbookFolder } from "../services/workingFolder.ts";

/**
 * Options for reading a workbook file.
 * @typedef {Object} ReadOptions
 * @property {WorkbookWorksheetName} [defaultWorksheetName] Default worksheet name to use when importing a CSV file.
 */
export type ReadOptions = {
	defaultWorksheetName?: WorkbookWorksheetName;
};

/**
 * Reads a workbook file (.xlsx or .csv) from a Microsoft Graph.
 * @param {DriveItemRef & Partial<DriveItem>} itemRef - Reference to the DriveItem to read from.
 * @param {ReadOptions} [options] - Options for reading, such as default worksheet name for CSV.
 * @returns {Promise<OpenRef>} Reference to the locally opened workbook.
 * @throws {InvalidArgumentError} If the file extension is not supported.
 */
export default async function readWorkbook(itemRef: DriveItemRef & Partial<DriveItem>, options: ReadOptions = {}): Promise<OpenRef> {
	const { defaultWorksheetName = defaultWorkbookWorksheetName } = options;

	const id = createOpenId();
	let name = itemRef.name;
	if (!name) {
		const item = await getDriveItem(itemRef);
		name = item.name ?? "";
	}

	const extension = extname(name).toLowerCase();

	const folder = await getWorkbookFolder(id);
	const targetFileName = join(folder, "0");
	const stream = await streamDriveItemContent(itemRef);

	if (extension === ".xlsx") {
		const fileStream = createWriteStream(targetFileName);
		await pipeline(stream, fileStream);
	} else if (extension === ".csv") {
		const xlsx = new ExcelJS.stream.xlsx.WorkbookWriter({ filename: targetFileName });
		const workbook = xlsx.addWorksheet(defaultWorksheetName);

		await new Promise<void>((resolve, reject) => {
			stream
				.pipe(parse({ headers: false }))
				.on("error", reject)
				.on("data", (row: unknown[]) => {
					workbook.addRow(row).commit();
				})
				.on("end", async () => {
					await xlsx.commit();
					resolve();
				});
		});
	} else {
		throw new InvalidArgumentError(`Unsupported file extension "${extension}".`);
	}

	return {
		id,
		itemRef,
	};
}
