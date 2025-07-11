import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import type { WorkbookWorksheetName } from "microsoft-graph/WorkbookWorksheet";
import writeWorkbook from "../src/tasks/writeWorkbook";
import { getLargeSet, getMemoryLimitMB } from "./shared";

(async () => {
	const memoryLimit = await getMemoryLimitMB();

	console.info(`Memory limit: ${memoryLimit.toFixed(2)} MB`);
	console.info("Creating large XLSX file...");

	const rows = getLargeSet();
	const itemName = generateTempFileName("xlsx");
	const itemPath = driveItemPath(itemName);
	const driveRef = getDefaultDriveRef();

	const uploadStart = Date.now();
	const item = await writeWorkbook(
		driveRef,
		itemPath,
		{ ["Sheet1" as WorkbookWorksheetName]: rows },
		{
			progress: ({ prepared, compressionRatio, written, preparedPerSecond, writtenPerSecond }) => {
				console.log(`[${new Date().toLocaleTimeString()}] Prepared ${prepared.toLocaleString()} (${preparedPerSecond.toLocaleString()}/sec), ${Math.round(compressionRatio * 100)}% compression, written: ${written.toLocaleString()} (${writtenPerSecond.toLocaleString()}/sec) cells`);
			},
		},
	);

	console.info(`Created XLSX: ${item.id} (${item.name}) at ${((item.size ?? 0) / 1024 / 1024).toLocaleString()} MB`);
	const totalSec = (Date.now() - uploadStart) / 1000;
	console.info(`Total runtime: ${totalSec.toFixed(2)} seconds`);
})();
