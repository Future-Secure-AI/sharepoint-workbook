import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { CellRef, ColumnComponent, ColumnNumber, RangeRef, RowComponent, RowNumber } from "../models/Reference.ts";
import type { WorksheetName } from "../models/Worksheet.ts";

const cellPattern = /^(?<col>[A-Z]{1,3})(?<row>\d{1,7})$/;
const cellRefPattern = /^(?:'(?<worksheet_quoted>[^']+)'|(?<worksheet_unquoted>[^'!]+))!(?<col>[A-Z]{1,3})(?<row>\d{1,7})$/;

export function parseCellReference(cell: CellRef): [worksheet: WorksheetName, col: ColumnNumber, row: RowNumber] {
	if (Array.isArray(cell)) {
		if (cell.length !== 2) {
			throw new Error(`Invalid cell reference array: ${cell}`);
		}
		const worksheet = cell[0] as WorksheetName;

		const match = cell[1].toString().match(cellPattern);
		if (!match || !match.groups) {
			throw new Error(`Invalid cell reference format: ${cell[1]}`);
		}

		const col = columnComponentToNumber(match.groups["col"] as ColumnComponent);
		const row = rowComponentToNumber(match.groups["row"] as RowComponent);
		return [worksheet, col, row];
	} else {
		const match = cell.toString().match(cellRefPattern);
		if (!match || !match.groups) {
			throw new Error(`Invalid cell reference format: ${cell}`);
		}

		// Use quoted or unquoted worksheet name, whichever matched
		const worksheet = (match.groups["worksheet_quoted"] ?? match.groups["worksheet_unquoted"]) as WorksheetName;
		const col = columnComponentToNumber(match.groups["col"] as ColumnComponent);
		const row = rowComponentToNumber(match.groups["row"] as RowComponent);

		return [worksheet, col, row];
	}
}
const rangePattern = /^(?:'(?<worksheet_quoted>[^']+)'|(?<worksheet_unquoted>[^'!]+))!(?<startCol>[A-Z]+)(?<startRow>\d+):(?<endCol>[A-Z]+)(?<endRow>\d+)$/;

/**
 * Converts a RangeRef to an array: [worksheet, startCol, startRow, endCol, endRow].
 * @param range RangeRef (array or string)
 * @param usedCols Number of columns in worksheet (for fallback)
 * @param usedRows Number of rows in worksheet (for fallback)
 * @returns [worksheet, startCol, startRow, endCol, endRow]
 */
export function parseRangeReference(range: RangeRef): [worksheet: WorksheetName, startCol: number | null, startRow: number | null, endCol: number | null, endRow: number | null] {
	if (Array.isArray(range)) {
		if (range.length !== 3) throw new Error(`Invalid range reference array: ${range}`);
		const [ws, start, end] = range;
		const parse = (val: string | number | undefined) => {
			if (typeof val === "string") {
				const m = val.match(/^([A-Z]+)?(\d+)?$/);
				const col = m?.[1];
				const row = m?.[2];
				return [col && /^[A-Z]+$/.test(col) ? (col as ColumnComponent) : undefined, row ? (row as RowComponent) : undefined] as [ColumnComponent?, RowComponent?];
			} else if (typeof val === "number") {
				return [undefined, val as RowComponent];
			}
			return [undefined, undefined];
		};
		const [startColRaw, startRowRaw] = parse(start ?? undefined);
		const [endColRaw, endRowRaw] = parse(end ?? undefined);
		const worksheet = ws as WorksheetName;
		const startColNum = typeof startColRaw === "string" && /^[A-Z]+$/.test(startColRaw) ? columnComponentToNumber(startColRaw as ColumnComponent) : null;
		const startRowNum = startRowRaw ? rowComponentToNumber(startRowRaw) : null;
		const endColNum = typeof endColRaw === "string" && /^[A-Z]+$/.test(endColRaw) ? columnComponentToNumber(endColRaw as ColumnComponent) : null;
		const endRowNum = endRowRaw ? rowComponentToNumber(endRowRaw) : null;
		if ((startColNum !== null && endColNum !== null && endColNum < startColNum) || (startRowNum !== null && endRowNum !== null && endRowNum < startRowNum)) {
			throw new InvalidArgumentError(`Range ends before it starts: [${worksheet}, ${startColNum}, ${startRowNum}, ${endColNum}, ${endRowNum}]`);
		}
		return [worksheet, startColNum, startRowNum, endColNum, endRowNum];
	} else if (typeof range === "string") {
		const match = range.match(rangePattern);
		if (!match?.groups) throw new Error(`Invalid range reference format: ${range}`);
		// Use quoted or unquoted worksheet name, whichever matched
		const worksheet = (match.groups["worksheet_quoted"] ?? match.groups["worksheet_unquoted"]) as WorksheetName;
		const startCol = match.groups["startCol"] ? columnComponentToNumber(match.groups["startCol"] as ColumnComponent) : null;
		const startRow = match.groups["startRow"] ? rowComponentToNumber(match.groups["startRow"] as RowComponent) : null;
		const endCol = match.groups["endCol"] ? columnComponentToNumber(match.groups["endCol"] as ColumnComponent) : null;
		const endRow = match.groups["endRow"] ? rowComponentToNumber(match.groups["endRow"] as RowComponent) : null;
		if ((startCol !== null && endCol !== null && endCol < startCol) || (startRow !== null && endRow !== null && endRow < startRow)) {
			throw new InvalidArgumentError(`Range ends before it starts: [${worksheet}, ${startCol}, ${startRow}, ${endCol}, ${endRow}]`);
		}
		return [worksheet, startCol, startRow, endCol, endRow];
	}
	throw new Error(`Invalid range reference: ${range}`);
}

export function columnComponentToNumber(column: ColumnComponent): ColumnNumber {
	let num = 0;
	for (let i = 0; i < column.length; i++) {
		num = num * 26 + (column.charCodeAt(i) - 65 + 1);
	}
	return num as ColumnNumber;
}

export function rowComponentToNumber(row: RowComponent): RowNumber {
	if (typeof row === "number") {
		return row;
	}
	const parsedRow = parseInt(row, 10);
	if (Number.isNaN(parsedRow)) {
		throw new Error(`Invalid row component: ${row}`);
	}
	return parsedRow as RowNumber;
}
