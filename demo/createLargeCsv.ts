import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import createWorkbook from "../src/tasks/createWorkbook";
import { getLargeSet } from "./shared";

(async () => {
    const rows = getLargeSet();
    const itemName = generateTempFileName("csv");
    const itemPath = driveItemPath(itemName);
    const driveRef = getDefaultDriveRef();
    const item = await createWorkbook(driveRef, itemPath, rows, {
        progress: (pct) => console.log(`Progress: ${Math.round(pct)}%`),
    });

    console.info(`Created item: ${item.id} (${item.name})`);
})();