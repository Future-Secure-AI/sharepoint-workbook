import { getDefaultDriveRef } from "microsoft-graph/drive";
import type { DriveItemPath } from "microsoft-graph/DriveItem";
import readWorkbookByPath from "../src/tasks/readWorkbookByPath";
import writeWorkbookByPath from "../src/tasks/writeWorkbookByPath";

const readFile = "/yellow_tripdata_2019-03.csv" as DriveItemPath;
const writeFile = "/yellow_tripdata_2019-03.xlsx" as DriveItemPath;

const start = Date.now();
const driveRef = getDefaultDriveRef();

console.info(`Reading file from SharePoint: ${readFile}`);
const ref = await readWorkbookByPath(driveRef, readFile, {
	progress: (bytes) => {
		console.info(`  Read ${(bytes / 1024 / 1024).toLocaleString()} MB`);
	},
});

console.info(`Writing file to SharePoint: ${writeFile}`);
await writeWorkbookByPath(ref, driveRef, writeFile, {
	ifExists: "replace",
	progress: (bytes) => {
		console.info(`  Written ${(bytes / 1024 / 1024).toLocaleString()} MB`);
	},
});

const elapsedMin = (Date.now() - start) / 1000 / 60;
console.info(`Done in ${elapsedMin.toFixed(2)} min(s)`);
