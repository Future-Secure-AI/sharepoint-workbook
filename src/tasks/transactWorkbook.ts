import AsposeCells from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import { iterateToArray } from "microsoft-graph/iteration";
import NotFoundError from "microsoft-graph/NotFoundError";
import picomatch from "picomatch";
import type { Cell, CellWrite } from "../models/Cell.ts";
import type { Handle } from "../models/Handle.ts";
import type { CellRef, RangeRef } from "../models/Reference.ts";
import type { RowWrite } from "../models/Row.ts";
import type { DeleteShift, InsertShift } from "../models/Shift.ts";
import type { WorksheetName } from "../models/Worksheet.ts";
import { normalizeCellWrite } from "../services/cell.ts";
import { parseCellReference, parseRangeReference } from "../services/reference.ts";
import { getLatestRevisionFilePath, getNextRevisionFilePath } from "../services/temporaryFile.ts";

type TransactContext = (operations: TransactOperations) => Promise<void>;

type TransactOperations = {
	listWorksheets: () => WorksheetName[];
	tryFindWorksheet: (search: string) => WorksheetName | null;
	findWorksheet: (search: string) => WorksheetName;

	readCells: (range: RangeRef) => Cell[][];
	updateCells: (origin: CellRef, cells: Iterable<RowWrite> | AsyncIterable<RowWrite>) => Promise<void>;
	insertCells: (origin: CellRef, shift: InsertShift, cells: Iterable<RowWrite> | AsyncIterable<RowWrite>) => Promise<void>;
	deleteCells: (range: RangeRef, shift: DeleteShift) => void;

	updateEachCell(range: RangeRef, write: CellWrite): void;
};

 const file = await getLatestRevisionFilePath(handle.localFilePath);
 let isDirty = false;
 const workbook = new AsposeCells.Workbook(file);

 // TODO: Named ranges, tables, pivot tables
 // TODO: Merging cells

 const listWorksheets = () => {
   const count = workbook.worksheets.count;
   const names: WorksheetName[] = [];
   for (let i = 0; i < count; i++) {
	 names.push(workbook.worksheets.get(i).name as WorksheetName);
   }
   return names;
 };

 const tryFindWorksheet = (search: string): WorksheetName | null => {
   const matcher = picomatch(search, { nocase: true });
   for (let i = 0; i < workbook.worksheets.count; i++) {
	 const ws = workbook.worksheets.get(i);
	 if (matcher(ws.name)) return ws.name as WorksheetName;
   }
   return null;
 };

 const findWorksheet = (search: string): WorksheetName => {
   const worksheet = tryFindWorksheet(search);
   if (!worksheet) throw new NotFoundError(`Worksheet not found for search: ${search}`);
   return worksheet;
 };

 const readCells = (range: RangeRef): Cell[][] => {
   let [worksheetName, startCol, startRow, endCol, endRow] = parseRangeReference(range);
   const worksheet = workbook.worksheets.get(workbook.worksheets.getSheetIndex(worksheetName));

   startCol = startCol ?? 0;
   startRow = startRow ?? 0;
   endCol = endCol ?? worksheet.cells.maxDataColumn;
   endRow = endRow ?? worksheet.cells.maxDataRow;

   const outputRows: Cell[][] = [];
   for (let r = startRow; r <= endRow; r++) {
	 const outputRow: Cell[] = [];
	 for (let c = startCol; c <= endCol; c++) {
	   const inputCell = worksheet.cells.get(r, c);
	   // fromExcelCell: convert Aspose cell to your Cell type
	   outputRow.push({
		 value: inputCell?.value,
		 // Add more fields as needed
	   } as Cell);
	 }
	 outputRows.push(outputRow);
   }
   return outputRows;
 };

 const updateCells = async (origin: CellRef, cells: Iterable<RowWrite> | AsyncIterable<RowWrite>): Promise<void> => {
   const [worksheetName, col, row] = parseCellReference(origin);
   const worksheet = workbook.worksheets.get(workbook.worksheets.getSheetIndex(worksheetName));

   let r = 0;
   for await (const inputRow of cells) {
	 let c = 0;
	 for (const inputCell of inputRow) {
	   const write = normalizeCellWrite(inputCell);
	   const cellObj = worksheet.cells.get(row + r, col + c);
	   if (write.value !== undefined) cellObj.value = write.value;
	   // Add more property assignments as needed
	   c++;
	 }
	 r++;
   }
   isDirty = true;
 };

 const insertCells = async (origin: CellRef, shift: InsertShift, cells: Iterable<RowWrite> | AsyncIterable<RowWrite>): Promise<void> => {
   const [worksheetName, col, row] = parseCellReference(origin);
   const worksheet = workbook.worksheets.get(workbook.worksheets.getSheetIndex(worksheetName));
   const rows = await iterateToArray(cells, (row) => row.map(normalizeCellWrite));

   if (shift === "Down") {
	 // Insert rows at 'row', shifting existing rows down
	 worksheet.cells.insertRows(row, rows.length);
	 rows.forEach((inputRow, rIdx) => {
	   inputRow.forEach((write, cIdx) => {
		 const cellObj = worksheet.cells.get(row + rIdx, col + cIdx);
		 if (write.value !== undefined) cellObj.value = write.value;
	   });
	 });
   } else if (shift === "Right") {
	 // Insert columns at 'col', shifting existing columns right
	 const maxCols = Math.max(0, ...rows.map((r) => r.length));
	 worksheet.cells.insertColumns(col, maxCols);
	 rows.forEach((inputRow, rIdx) => {
	   inputRow.forEach((write, cIdx) => {
		 const cellObj = worksheet.cells.get(row + rIdx, col + cIdx);
		 if (write.value !== undefined) cellObj.value = write.value;
	   });
	 });
   } else {
	 throw new InvalidArgumentError(`Unsupported shift: ${shift}`);
   }
   isDirty = true;
 };

 const deleteCells = (range: RangeRef, shift: DeleteShift) => {
   let [worksheetName, startCol, startRow, endCol, endRow] = parseRangeReference(range);
   const worksheet = workbook.worksheets.get(workbook.worksheets.getSheetIndex(worksheetName));

   startCol = startCol ?? 0;
   startRow = startRow ?? 0;
   endCol = endCol ?? worksheet.cells.maxDataColumn;
   endRow = endRow ?? worksheet.cells.maxDataRow;

   if (shift === "Up") {
	 const numRows = endRow - startRow + 1;
	 worksheet.cells.deleteRows(startRow, numRows);
   } else if (shift === "Left") {
	 const numCols = endCol - startCol + 1;
	 worksheet.cells.deleteColumns(startCol, numCols);
   } else {
	 throw new InvalidArgumentError(`Unsupported shift: ${shift}`);
   }
   isDirty = true;
 };

 const updateEachCell = (range: RangeRef, write: CellWrite): void => {
   let [worksheetName, startCol, startRow, endCol, endRow] = parseRangeReference(range);
   const worksheet = workbook.worksheets.get(workbook.worksheets.getSheetIndex(worksheetName));

   startCol = startCol ?? 0;
   startRow = startRow ?? 0;
   endCol = endCol ?? worksheet.cells.maxDataColumn;
   endRow = endRow ?? worksheet.cells.maxDataRow;

   for (let r = startRow; r <= endRow; r++) {
	 for (let c = startCol; c <= endCol; c++) {
	   const cellObj = worksheet.cells.get(r, c);
	   if (write.value !== undefined) cellObj.value = write.value;
	 }
   }
   isDirty = true;
 };

 const operations: TransactOperations = {
   listWorksheets,
   tryFindWorksheet,
   findWorksheet,
   readCells,
   updateCells,
   insertCells,
   deleteCells,
   updateEachCell,
 };

 await context(operations);

 if (isDirty) {
   const nextFile = await getNextRevisionFilePath(handle.localFilePath);
   workbook.save(nextFile);
 }
}
