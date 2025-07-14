import { getDefaultDriveRef } from "microsoft-graph/drive";
import type { DriveItemPath } from "microsoft-graph/DriveItem";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import optimizeWorkbook from "../src/tasks/optimizeWorkbook";
import readWorkbookByPath from "../src/tasks/readWorkbookByPath";
import writeWorkbookByPath from "../src/tasks/writeWorkbookByPath";

const readFile = "/yellow_tripdata_2019-03.csv" as DriveItemPath;
const writeFile = driveItemPath(generateTempFileName("xlsx")) as DriveItemPath;

const start = Date.now();
const driveRef = getDefaultDriveRef();

console.info(`Reading input CSV from SharePoint...`);
const ref = await readWorkbookByPath(driveRef, readFile, {
	progress: (bytes) => {
		console.info(`  Read ${Math.round(bytes / 1024 / 1024).toLocaleString()} MB...`);
	},
});

console.info(`Optimizing workbook...`);
const ratio = await optimizeWorkbook(ref);
console.info(`  Reduced file size by ${Math.round((1 - ratio) * 100)}%`);

console.info(`Writing output XLSX file to SharePoint...`);
await writeWorkbookByPath(ref, driveRef, writeFile, {
	ifExists: "replace",
	maxChunkSize: 15 * 1024 * 1024, // Best speed with 60MB chunks (max), but using smaller to get more frequent progress updates
	progress: (bytes) => {
		console.info(`  Written ${Math.round(bytes / 1024 / 1024).toLocaleString()} MB...`);
	},
});

const elapsedMin = (Date.now() - start) / 1000 / 60;
console.info(`Done in ${elapsedMin.toFixed(2)} min(s) as ${writeFile}.`);
