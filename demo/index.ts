import { getDefaultDriveRef } from "microsoft-graph/drive";
import type { DriveItemPath } from "microsoft-graph/DriveItem";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import optimizeWorkbook from "../src/tasks/optimizeWorkbook";
import readWorkbookByPath from "../src/tasks/readWorkbookByPath";
import transactWorkbook from "../src/tasks/transactWorkbook";
import writeWorkbookByPath from "../src/tasks/writeWorkbookByPath";
const readFile = "/yellow_tripdata_2019-03.csv" as DriveItemPath;
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

/*
 * Do some work on the workbook. This sample will delete some columns and format the header row.
 * You can do anything you want here, like adding formulas, formatting, etc. Just remember that
 * up until this point the file hasn't needed to be in memory. `transact` requires sufficient memory
 * to hold the whole workbook.
 */
console.info(`Modifying workbook...`);
await transactWorkbook(handle, async ({ findWorksheet, deleteCells, updateEachCell }) => {
	const sheet = findWorksheet("*");
	deleteCells([sheet, "D", "R"], "Left");
	updateEachCell([sheet, 1, 1], {
		fontBold: true,
	});
});

/*
 * Recompress the workbook to reduce it's size. Effectively you're spending CPU to save upload time, storage space, and perhaps getting the
 * file under the 100MB SharePoint web size limit. This is completely optional, and there is an option to set the compression level. I don't
 * recommend anything but 6, as it can take a long time to compress and doesn't save much more space. But if every byte counts, go for 9.
 */
console.info(`Optimizing workbook... (may take a while)`);
const ratio = await optimizeWorkbook(handle, {
	compressionLevel: 6,
});
console.info(`  Reduced file size by ${Math.round((1 - ratio) * 100)}%`);

/*
 * Write the workbook back to SharePoint in a location of your choosing. Only writing XLSX is supported.
 */
console.info(`Writing output XLSX '${writeFile}' to SharePoint...`);
await writeWorkbookByPath(handle, driveRef, writeFile, {
	ifExists: "replace",
	maxChunkSize: 15 * 1024 * 1024, // Best speed with 60MB chunks (max), but for this demo I'm using a smaller value to get more frequent progress updates.
	progress: (bytes) => {
		console.info(`  Written ${formatBytes(bytes)}...`);
	},
});

const elapsedMin = (Date.now() - start) / 1000 / 60;
console.info(`Done in ${elapsedMin.toFixed(1)} min(s).`);

function formatBytes(bytes: number): string {
	return `${Math.round(bytes / 1024 / 1024).toLocaleString()} MB`;
}
