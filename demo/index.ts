/** biome-ignore-all lint/correctness/noUnusedFunctionParameters: We want these included for demonstration purposes. */

import type { ColumnName } from "microsoft-graph/Column";
import { getDefaultDriveRef } from "microsoft-graph/drive";
import type { DriveItemPath } from "microsoft-graph/DriveItem";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import console from "node:console";
import addPivotTableData from "../src/tasks/addPivotTableData.ts";
import addPivotTableRow from "../src/tasks/addPivotTableRow.ts";
import createPivotTable from "../src/tasks/createPivotTable.ts";
import createWorksheet from "../src/tasks/createWorksheet.ts";
import findWorksheet from "../src/tasks/findWorksheet.ts";
import openWorkbook from "../src/tasks/openWorkbook.ts";
import saveWorkbookAs from "../src/tasks/saveWorkbookAs.ts";
import updateEachCell from "../src/tasks/updateEachCell.ts";

const readFile = "/700MB.csv.gz" as DriveItemPath;
const writeFile = driveItemPath(generateTempFileName("xlsb")) as DriveItemPath;

const start = Date.now();
const driveRef = getDefaultDriveRef();

/*
 * Read an input file from Sharepoint. It could be any file that is:
 * - No more than 250GB
 * - No more than 4x amount of available server memory
 * - No more than the configured Node memory limit (default 4GB) less what's already used
 * - Is a supported file type https://docs.aspose.com/cells/cpp/supported-file-formats/
 */
console.info(`Reading input '${readFile}' from SharePoint...`);
const workbook = await openWorkbook(driveRef, readFile, {
	progress: (bytes) => {
		console.info(`  Read ${formatBytes(bytes)}...`);
	},
});

/*
 * Apply some formatting. In this case I'm making the header bold
 */
console.info(`Formatting header row...`);
const dataSheet = findWorksheet(workbook, "*"); // First worksheet
updateEachCell(dataSheet, "1", {
	fontBold: true,
});

/*
 * Create a pivot table in a new worksheet.
 */
console.info(`Creating pivot table...`);
const pivotSheet = createWorksheet(workbook, "pivot");
const pivotTable = createPivotTable(pivotSheet, dataSheet, "A:R");
addPivotTableRow(pivotTable, "VendorID" as ColumnName);
addPivotTableData(pivotTable, "fare_amount" as ColumnName);

/*
 * Write the workbook back to SharePoint in a location of your choosing.
 */
console.info(`Writing output '${writeFile}' to SharePoint...`);
await saveWorkbookAs(workbook, driveRef, writeFile, {
	ifExists: "replace",
	maxChunkSize: 30 * 1024 * 1024, // Best speed with 60MB chunks (max), but for this demo I'm using a smaller value to get more frequent progress updates.
	progress: (bytes) => {
		console.info(`  Written ${formatBytes(bytes)}...`);
	},
});

const elapsedMin = (Date.now() - start) / 1000 / 60;
console.info(`Done in ${elapsedMin.toFixed(1)} min(s).`); // Done in 0.8 min(s).

function formatBytes(bytes: number): string {
	return `${Math.round(bytes / 1024 / 1024).toLocaleString()} MB`;
}
