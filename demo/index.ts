/** biome-ignore-all lint/correctness/noUnusedFunctionParameters: We want these included for demonstration purposes. */

import { getDefaultDriveRef } from "microsoft-graph/drive";
import type { DriveItemPath } from "microsoft-graph/DriveItem";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import console from "node:console";
import { stat } from "node:fs/promises";
import readWorkbookByPath from "../src/tasks/readWorkbookByPath";
import writeWorkbookByPath from "../src/tasks/writeWorkbookByPath";

const readFile = "/700MB.csv" as DriveItemPath;
const writeFile = driveItemPath(generateTempFileName("xlsx")) as DriveItemPath;

const start = Date.now();
const driveRef = getDefaultDriveRef();

/*
 * Read an input file from Sharepoint. It could the CSV or XLSX, and it could be any size up to 250GB. This sample is about 730MB.
 */
console.info(`Reading input CSV '${readFile}' from SharePoint...`);
const handle = await readWorkbookByPath(driveRef, readFile, {
	progress: (bytes) => {
		console.info(`  Read ${formatBytes(bytes)}...`);
	},
});

const { size } = await stat(handle.localFilePath);
console.info(handle.localFilePath, formatBytes(size));
console.info(`Input read and optimized down to ${formatBytes(size)}.`);

/*
 * Optionally filter out some columns or rows.
 */
// console.info(`Filtering workbook...`);
// await filterWorkbook(handle, {
// 	skipRows: 0,
// 	columnFilter: (header, index) => header === "tpep_dropoff_datetime" || header === "tolls_amount",
// 	rowFilter: (cells) => true,
// 	progress: (rows) => {
// 		console.info(`  Processed ${rows.toLocaleString()} rows...`);
// 	},
// });

/*
 * Do some work on the workbook. This sample just formats the header rows, but you can do anything you want here, like adding
 * formulas, formatting, etc. Just remember that up until this point the file hasn't needed to be in memory. `transact` requires
 * sufficient memory to hold the whole workbook.
 */
// console.info(`Formatting workbook...`);
// await transactWorkbook(handle, async ({ findWorksheet, updateEachCell }) => {
// 	const sheet = findWorksheet("*");
// 	updateEachCell([sheet, 1, 1], {
// 		fontBold: true,
// 	});
// });

/*
 * Write the workbook back to SharePoint in a location of your choosing. Only writing XLSX is supported.
 */
console.info(`Writing output XLSX '${writeFile}' to SharePoint...`);
await writeWorkbookByPath(handle, driveRef, writeFile, {
	ifExists: "replace",
	maxChunkSize: 30 * 1024 * 1024, // Best speed with 60MB chunks (max), but for this demo I'm using a smaller value to get more frequent progress updates.
	progress: (bytes) => {
		console.info(`  Written ${formatBytes(bytes)}...`);
	},
});

const elapsedMin = (Date.now() - start) / 1000 / 60;
console.info(`Done in ${elapsedMin.toFixed(1)} min(s).`);

function formatBytes(bytes: number): string {
	return `${Math.round(bytes / 1024 / 1024).toLocaleString()} MB`;
}
