import { type CsvFormatterStream, format, type Row } from "@fast-csv/format";
import type { DriveItem } from "@microsoft/microsoft-graph-types";
import ExcelJS from "exceljs";
import type { UsedAddress } from "microsoft-graph/Address";
import type { Cell, CellScope } from "microsoft-graph/Cell";
import createDriveItemContent from "microsoft-graph/createDriveItemContent";
import type { DriveItemPath, DriveItemRef } from "microsoft-graph/DriveItem";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { WorkbookWorksheetName } from "microsoft-graph/WorkbookWorksheet";
import { defaultWorkbookWorksheetName } from "microsoft-graph/workbookWorksheet";
import { extname } from "node:path";
import { Readable } from "node:stream";

export type ReadOptions = {
    address?: UsedAddress;
    scope?: Partial<CellScope>;
};

export type WriterOptions = {
    sheetName?: WorkbookWorksheetName;
};

export async function createWorkbook(parentRef: DriveItemRef, itemPath: DriveItemPath, rows: Iterable<Partial<Cell>[]> | AsyncIterable<Partial<Cell>[]>, options: WriterOptions = {}): Promise<DriveItem & DriveItemRef> {
    const extension = extname(itemPath);

    if (extension === ".csv") {
        const writer = format({ headers: false });
        for await (const row of rows) {
            writer.write(row.map((cell) => cell.value ?? ""));
        }
        writer.end();

        const buffer = await writeToBuffer(writer);
        const stream = Readable.from(buffer);
        return await createDriveItemContent(parentRef, itemPath, stream, buffer.length, "fail");
    } else if (extension === ".xlsx") {
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet(options.sheetName || defaultWorkbookWorksheetName);

        for await (const row of rows) {
            worksheet.addRow(row.map((cell) => cell.value ?? ""));
        }

        const buffer = Buffer.from(await workbook.xlsx.writeBuffer());
        const stream = Readable.from([buffer]);
        return await createDriveItemContent(parentRef, itemPath, stream, buffer.byteLength, "fail");
    }

    throw new InvalidArgumentError(`Unsupported file extension: ${extension}. Supported extensions are .csv and .xlsx.`);
}

async function writeToBuffer(writer: CsvFormatterStream<Row, Row>): Promise<Buffer> {
    const chunks: Buffer[] = [];
    for await (const chunk of writer) {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }
    const buff = Buffer.concat(chunks);
    return buff;
}

// export async function writeWorkbookCells(itemRef: DriveItemRef, rows: Iterable<Partial<Cell>[]> | AsyncIterable<Partial<Cell>[]>, options: WriteOptions = {}): Promise<void> { }

// export async function* readWorkbookCells(itemRef: DriveItemRef, options: ReadOptions = {}): AsyncIterable<IteratedRow> { }
