import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import createWorkbook from "../src/tasks/createWorkbook";
import { getLargeSet, getMemoryLimitMB } from "./shared";

(async () => {
    const memoryLimit = await getMemoryLimitMB();

    console.info(`Memory limit: ${memoryLimit.toFixed(2)} MB`);
    console.info("Creating large XLSX file...")

    const rows = getLargeSet();
    const itemName = generateTempFileName("xlsx");
    const itemPath = driveItemPath(itemName);
    const driveRef = getDefaultDriveRef();
    const uploadStart = Date.now();
    const item = await createWorkbook(driveRef, itemPath, rows, {
        encodeProgress: (rowCount) => console.log(`Encoded ${rowCount.toLocaleString()} rows`),
        uploadProgress: (rowCount, pct) => {
            const elapsedSec = (Date.now() - uploadStart) / 1000;
            const rowsPerSec = elapsedSec > 0 ? rowCount / elapsedSec : 0;
            console.log(`Uploaded ${Math.round(pct * 100) / 100}% ${rowCount.toLocaleString()} rows (${rowsPerSec.toFixed(2)} rows/sec)`);
        },
    });

    console.info(`Created item: ${item.id} (${item.name})`);
})();