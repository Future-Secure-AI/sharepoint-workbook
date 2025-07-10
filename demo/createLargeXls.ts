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
        progress: (preparedCount, uploadedCount, preparedPerSecond, writtenPerSecond) => {
            console.log(`Progress: ` +
                `${preparedCount.toLocaleString()} prepared\t ` +
                `${uploadedCount.toLocaleString()} uploaded\t ` +
                `(${preparedPerSecond.toLocaleString()} prepared/sec, ${writtenPerSecond.toLocaleString()} uploaded/sec)`);
        },
    });

    console.info(`Created XLSX: ${item.id} (${item.name}) at ${item.size ?? 0 / 1024 / 1024} MB`);
    const totalSec = (Date.now() - uploadStart) / 1000;
    console.info(`Total runtime: ${totalSec.toFixed(2)} seconds`);
})();