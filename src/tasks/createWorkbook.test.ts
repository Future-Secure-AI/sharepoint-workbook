
import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import tryDeleteDriveItem from "microsoft-graph/tryDeleteDriveItem";
import { describe, expect, it } from "vitest";
import { getLargeSet } from "../../demo/shared.ts";
import createWorkbook from "./createWorkbook.ts";

function getSmallSet() {
    return [
        [{ value: "A" }, { value: "B" }, { value: "C" }],
        [{ value: "D" }, { value: "E" }, { value: "F" }],
        [{ value: "G" }, { value: "H" }, { value: "I" }],
    ]
}


describe("createWorkbook", { timeout: 15 * 60 * 1000 }, () => {
    it("creates small CSV file", async () => {
        const rows = getSmallSet();
        const itemName = generateTempFileName("csv");
        const itemPath = driveItemPath(itemName);
        const driveRef = getDefaultDriveRef();
        const item = await createWorkbook(driveRef, itemPath, rows);
        expect(item).toBeTruthy();
        expect(item.name).toBe(itemName);
        expect(item.size).toBeGreaterThan(0);
        await tryDeleteDriveItem(item);
    });

    it("creates small XLSX file", async () => {
        const rows = getSmallSet();
        const itemName = generateTempFileName("xlsx");
        const itemPath = driveItemPath(itemName);
        const driveRef = getDefaultDriveRef();
        const item = await createWorkbook(driveRef, itemPath, rows);
        expect(item).toBeTruthy();
        expect(item.name).toBe(itemName);
        expect(item.size).toBeGreaterThan(0);
        await tryDeleteDriveItem(item);
    });

    it("throws for unsupported file extension", async () => {
        const rows = getSmallSet();
        const itemName = generateTempFileName("txt");
        const itemPath = driveItemPath(itemName);
        const driveRef = getDefaultDriveRef();
        await expect(createWorkbook(driveRef, itemPath, rows)).rejects.toThrow(
            /Unsupported file extension/
        );
    });

    it("creates large CSV file from NYC dataset subset", async () => {
        const rows = getLargeSet();
        const itemName = generateTempFileName("csv");
        const itemPath = driveItemPath(itemName);
        const driveRef = getDefaultDriveRef();
        const item = await createWorkbook(driveRef, itemPath, rows, {
            progress: (pct) => console.log(`Progress: ${Math.round(pct)}%`),
        });
        expect(item).toBeTruthy();
        expect(item.name).toBe(itemName);
        expect(item.size).toBeGreaterThan(0);
        // await tryDeleteDriveItem(item);
    });

    it("creates large XLSX file from NYC dataset subset", async () => {
        const rows = getLargeSet();
        const itemName = generateTempFileName("xlsx");
        const itemPath = driveItemPath(itemName);
        const driveRef = getDefaultDriveRef();
        const item = await createWorkbook(driveRef, itemPath, rows, {
            progress: (pct) => console.log(`Progress: ${Math.round(pct)}%`),
        });
        expect(item).toBeTruthy();
        expect(item.name).toBe(itemName);
        expect(item.size).toBeGreaterThan(0);
        await tryDeleteDriveItem(item);
    });
});
