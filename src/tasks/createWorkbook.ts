/**
 * Copy a drive item.
 * @module createWorkbook
 * @category Tasks
 */

import { format } from "@fast-csv/format";
import type { DriveItem } from "@microsoft/microsoft-graph-types";
import ExcelJS from "exceljs";
import type { Cell } from "microsoft-graph/Cell";
import createDriveItemContent from "microsoft-graph/createDriveItemContent";
import type { DriveRef } from "microsoft-graph/Drive";
import type { DriveItemPath, DriveItemRef } from "microsoft-graph/DriveItem";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { WorkbookWorksheetName } from "microsoft-graph/WorkbookWorksheet";
import { defaultWorkbookWorksheetName } from "microsoft-graph/workbookWorksheet";
import { extname } from "node:path";
import { Readable } from "node:stream";
import { csvToBuffer } from "../services/csvFormatter.ts";
import { appendRow, xlsxToBuffer } from "../services/excelJs.ts";

/**
 * Options for creating a new workbook file.
 * @property {WorkbookWorksheetName} [sheetName] Name of the worksheet to create (for .xlsx files).
 * @property {"fail" | "replace" | "rename"} [conflictResolution] How to resolve name conflicts when uploading the file.
 */
export type CreateOptions = {
    sheetName?: WorkbookWorksheetName;
    conflictBehavior?: "fail" | "replace" | "rename";
    chunkSize?: number;
    progress?: (pct: number) => void;
};

/**
 * Creates a new workbook (.csv or .xlsx) in the specified parent location with the provided rows.
 * @param {DriveRef | DriveItemRef} parentRef Reference to the parent drive or item where the file will be created.
 * @param {DriveItemPath} itemPath Path (including filename and extension) for the new workbook.
 * @param {Iterable<Partial<Cell>[]> | AsyncIterable<Partial<Cell>[]>} rows Iterable or async iterable of row arrays, each containing partial Cell objects.
 * @param {CreateOptions} [options] Options for sheet name and conflict resolution.
 * @returns {Promise<DriveItem & DriveItemRef>} Created DriveItem with reference.
 * @throws {InvalidArgumentError} If the file extension is not supported.
 */
export default async function createWorkbook(parentRef: DriveRef | DriveItemRef, itemPath: DriveItemPath, rows: Iterable<Partial<Cell>[]> | AsyncIterable<Partial<Cell>[]>, { sheetName, conflictBehavior, chunkSize, progress }: CreateOptions = {}): Promise<DriveItem & DriveItemRef> {
    let buffer: Buffer;
    const extension = extname(itemPath);
    if (extension === ".csv") {
        const csv = format({ headers: false });
        for await (const row of rows) {
            csv.write(row.map((cell) => cell.value ?? ""));
        }
        csv.end();

        buffer = await csvToBuffer(csv);
    } else if (extension === ".xlsx") {
        const xls = new ExcelJS.Workbook();
        const worksheet = xls.addWorksheet(sheetName ?? defaultWorkbookWorksheetName);

        for await (const row of rows) {
            appendRow(worksheet, row);
        }

        buffer = await xlsxToBuffer(xls);
    } else {
        throw new InvalidArgumentError(`Unsupported file extension: ${extension}. Supported extensions are .csv and .xlsx.`);
    }

    const stream = Readable.from(buffer);

    return await createDriveItemContent(parentRef, itemPath, stream, buffer.byteLength, {
        conflictBehavior,
        chunkSize,
        progress,
    });
};

