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

/*
 * Read an input file from Sharepoint. It could the CSV or XLSX, and it could be any size up to 250GB.
 */
console.info(`Reading input CSV from SharePoint...`);
const ref = await readWorkbookByPath(driveRef, readFile, {
	progress: (bytes) => {
		console.info(`  Read ${Math.round(bytes / 1024 / 1024).toLocaleString()} MB...`);
	},
});

/*
 * Recompress the workbook to reduce it's size. Effectively you're spending CPU to save upload time, storage space, and perhaps getting the
 * file under the 100MB SharePoint web size limit. This is completely optional, and there is an option to set the compression level. I don't
 * recommend going above 6, as it can take a long time to compress and doesn't save much more space. But if every byte counts, you can go up to 9.
 */
console.info(`Optimizing workbook...`);
const ratio = await optimizeWorkbook(ref, { compressionLevel: 6 });
console.info(`  Reduced file size by ${Math.round((1 - ratio) * 100)}%`);

/*
 * Write the workbook back to SharePoint in a location of your choosing.
 */
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
