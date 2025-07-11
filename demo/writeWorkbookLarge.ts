import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import type { WriteWorksheet } from "../src/models/Worksheet";
import writeWorkbook from "../src/tasks/writeWorkbook";
import { getLargeSet, getMemoryLimitMB } from "./shared";

const progress = ({ prepared, compressionRatio, written, preparedPerSecond, writtenPerSecond }: import("/Users/ben.thompson/Library/CloudStorage/OneDrive-FutureSecureAI/Documents/GitHub/sharepoint-workbook/src/tasks/writeWorkbook").WriteProgress): void => {
	console.log(`[${new Date().toLocaleTimeString()}] Prepared ${prepared.toLocaleString()} (${preparedPerSecond.toLocaleString()}/sec), ${Math.round(compressionRatio * 100)}% compression, written: ${written.toLocaleString()} (${writtenPerSecond.toLocaleString()}/sec) cells`);
};

(async () => {
	const memoryLimit = await getMemoryLimitMB();

	console.info(`Memory limit: ${memoryLimit.toFixed(2)} MB`);
	console.info("Creating large XLSX file...");

	const rows = getLargeSet();
	const itemName = generateTempFileName("xlsx");
	const itemPath = driveItemPath(itemName);
	const driveRef = getDefaultDriveRef();

	const uploadStart = Date.now();
	const worksheets = [
		{
			name: "Sheet1",
			rows,
		} satisfies WriteWorksheet,
	];
	const item = await writeWorkbook(driveRef, itemPath, worksheets, { progress });

	console.info(`Created XLSX: ${item.id} (${item.name}) at ${((item.size ?? 0) / 1024 / 1024).toLocaleString()} MB`);
	const totalSec = (Date.now() - uploadStart) / 1000;
	console.info(`Total runtime: ${totalSec.toFixed(2)} seconds`);
})();
