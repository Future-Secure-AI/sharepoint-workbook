import type { WorksheetName } from "./Worksheet.ts";

type Column = `${Uppercase<string>}`;
type Row = `${number}`;
type Cell = `${Column}${Row}`;

export type ColumnRef = `${WorksheetName}!${Column}`;
export type RowRef = `${WorksheetName}!${Row}`;
export type CellRef = `${WorksheetName}!${Cell}`;

export type ColumnRangeRef = `${WorksheetName}!${Column}:${Column}`;
export type RowRangeRef = `${WorksheetName}!${Row}:${Row}`;
export type CellRangeRef = `${WorksheetName}!${Cell}:${Cell}`;

export type PartialColumnRangeRef = `${WorksheetName}!${Column}:` | `${WorksheetName}!:${Column}`;
export type PartialRowRangeRef = `${WorksheetName}!${Row}:` | `${WorksheetName}!:${Row}`;
export type PartialCellRangeRef = `${WorksheetName}!${Cell}:` | `${WorksheetName}!:${Cell}`;
export type PartialWorksheetRangeRef = `${WorksheetName}!:`;

export type Ref = ColumnRef | RowRef | CellRef | ColumnRangeRef | RowRangeRef | CellRangeRef;
export type PartialRef = Ref | PartialColumnRangeRef | PartialRowRangeRef | PartialCellRangeRef | PartialWorksheetRangeRef;
export type ParsedRef = [worksheet: WorksheetName, startCol: Column, startRow: Row, endCol: Column, endRow: Row];
