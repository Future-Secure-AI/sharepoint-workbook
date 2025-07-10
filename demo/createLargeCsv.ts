import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import createWorkbook from "../src/tasks/createWorkbook";
import { getLargeSet, getMemoryLimitMB } from "./shared";

(async () => {
    const memoryLimit = await getMemoryLimitMB();

    console.info(`Memory limit: ${memoryLimit.toFixed(2)} MB`);
    console.info("Creating large CSV file...")

    const rows = getLargeSet();
    const itemName = generateTempFileName("csv");
    const itemPath = driveItemPath(itemName);
    const driveRef = getDefaultDriveRef();
    const uploadStart = Date.now();
    let lastUploadCount = 0;
    let lastUploadTime = uploadStart;
    let lastEncodeCount = 0;
    const item = await createWorkbook(driveRef, itemPath, rows, {
        encodeProgress: (rowCount) => {
            if (rowCount - lastEncodeCount >= 100000) {
                console.log(`Encoded ${rowCount.toLocaleString()} rows`);
                lastEncodeCount = rowCount;
            }
        },
        uploadProgress: (rowCount, pct) => {
            const now = Date.now();
            const elapsedSec = (now - lastUploadTime) / 1000;
            const rowsSinceLast = rowCount - lastUploadCount;
            const rowsPerSec = elapsedSec > 0 ? rowsSinceLast / elapsedSec : 0;
            console.log(`Uploaded ${Math.round(pct * 100) / 100}% ${rowCount.toLocaleString()} rows (${rowsPerSec.toFixed(2)} rows/sec)`);
            lastUploadCount = rowCount;
            lastUploadTime = now;
        },
    });

    console.info(`Created CSV: ${item.id} (${item.name}) at ${item.size ?? 0 / 1024 / 1024} MB`);
    const totalSec = (Date.now() - uploadStart) / 1000;
    console.info(`Total runtime: ${totalSec.toFixed(2)} seconds`);
})();