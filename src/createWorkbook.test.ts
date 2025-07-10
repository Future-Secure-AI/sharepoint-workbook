import type { Cell } from "microsoft-graph/Cell";
import { getDefaultDriveRef } from "microsoft-graph/drive";
import { driveItemPath } from "microsoft-graph/driveItem";
import { generateTempFileName } from "microsoft-graph/temporaryFiles";
import tryDeleteDriveItem from "microsoft-graph/tryDeleteDriveItem";
import { describe, expect, it } from "vitest";
import { createWorkbook } from "./createWorkbook";

const rows: Partial<Cell>[][] = [
    [{ value: "A" }, { value: "B" }, { value: "C" }],
    [{ value: "D" }, { value: "E" }, { value: "F" }],
    [{ value: "G" }, { value: "H" }, { value: "I" }],
];

describe("createWorkbook", () => {
    it("creates a CSV file with correct values", async () => {
        const itemName = generateTempFileName("csv");
        const itemPath = driveItemPath(itemName);

        const driveRef = getDefaultDriveRef();
        const item = await createWorkbook(driveRef, itemPath, rows);
        expect(item).toBeTruthy();
        expect(item.name).toBe(itemName);
        expect(item.size).toBeGreaterThan(0);
        await tryDeleteDriveItem(item);
    });
});
