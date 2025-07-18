/** biome-ignore-all lint/complexity/useLiteralKeys: Impossible to avoid with RegEx */

import type { Worksheet } from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { ColumnNumber } from "../models/Column.ts";
import type { CellRef, ColumnRef, RangeRef, Ref, RowRef } from "../models/Reference.ts";
import type { RowNumber } from "../models/Row.ts";

// Matches cell refs like A1, Z99, etc.
const cellPattern = /^(?<col>[A-Z]{1,3})(?<row>\d{1,7})$/;
// Matches range refs like A1:C3
const rangePattern = /^(?<startCol>[A-Z]+)?(?<startRow>\d+)?:(?<endCol>[A-Z]+)?(?<endRow>\d+)?$/;

/**
 * Parses a cell reference (e.g., "A1") into [col, row] numbers.
 */
export function parseCellReference(cell: CellRef): [ColumnNumber, RowNumber] {
	const match = cell.toString().match(cellPattern)?.groups;
	if (!match) throw new Error(`Invalid cell reference format: '${cell}'`);
	return [columnComponentToNumber(match["col"] as ColumnRef), rowComponentToNumber(match["row"] as RowRef)];
}

/**
 * Converts a RangeRef to an array: [startCol, startRow, endCol, endRow].
 * @param range RangeRef (array or string)
 * @returns [startCol, startRow, endCol, endRow]
 */
export function parseRangeReference(range: RangeRef): [number | null, number | null, number | null, number | null] {
	if (!Array.isArray(range)) {
		const match = range.match(rangePattern)?.groups;
		if (!match) throw new Error(`Invalid range reference format: ${range}`);
		const startCol = match["startCol"] ? columnComponentToNumber(match["startCol"] as ColumnRef) : null;
		const startRow = match["startRow"] ? rowComponentToNumber(match["startRow"] as RowRef) : null;
		const endCol = match["endCol"] ? columnComponentToNumber(match["endCol"] as ColumnRef) : null;
		const endRow = match["endRow"] ? rowComponentToNumber(match["endRow"] as RowRef) : null;
		if ((startCol !== null && endCol !== null && endCol < startCol) || (startRow !== null && endRow !== null && endRow < startRow)) {
			throw new InvalidArgumentError(`Range ends before it starts: [${startCol}, ${startRow}, ${endCol}, ${endRow}]`);
		}
		return [startCol, startRow, endCol, endRow];
	}

	if (range.length !== 2) throw new Error(`Invalid range reference array: ${range}`);
	const [start, end] = range;

	// Row-only range: [number, number] or [string, string] with both numeric
	const isRowOnly = (v: unknown) => typeof v === "number" || (typeof v === "string" && /^\d+$/.test(v));
	// Column-only range: [string, string] with both single uppercase letters
	const isColOnly = (v: unknown) => typeof v === "string" && /^[A-Z]$/.test(v);

	if (isRowOnly(start) && isRowOnly(end)) {
		const startRow = rowComponentToNumber(start as RowRef);
		const endRow = rowComponentToNumber(end as RowRef);
		if (endRow < startRow) throw new InvalidArgumentError(`Range ends before it starts: [null, ${startRow}, null, ${endRow}]`);
		return [null, startRow, null, endRow];
	}
	if (isColOnly(start) && isColOnly(end)) {
		const startCol = columnComponentToNumber(start as ColumnRef);
		const endCol = columnComponentToNumber(end as ColumnRef);
		if (endCol < startCol) throw new InvalidArgumentError(`Range ends before it starts: [${startCol}, null, ${endCol}, null]`);
		return [startCol, null, endCol, null];
	}
	// Column-row range: [col, row] (e.g., ["A", 3] or ["A", "3"])
	if (isColOnly(start) && isRowOnly(end)) {
		const startCol = columnComponentToNumber(start as ColumnRef);
		const endRow = rowComponentToNumber(end as RowRef);
		return [startCol, null, null, endRow];
	}
	// Row-column range: [row, col] (e.g., [1, "C"] or ["1", "C"])
	if (isRowOnly(start) && isColOnly(end)) {
		const startRow = rowComponentToNumber(start as RowRef);
		const endCol = columnComponentToNumber(end as ColumnRef);
		return [null, startRow, endCol, null];
	}

	// Otherwise, treat as cell references
	const parseCell = (ref: Ref | null) => {
		if (ref == null) return [null, null];
		try {
			return parseCellReference(ref as CellRef);
		} catch {
			return [null, null];
		}
	};
	const [startCol, startRow] = parseCell(start);
	const [endCol, endRow] = parseCell(end);
	if ((startCol !== null && endCol !== null && endCol < startCol) || (startRow !== null && endRow !== null && endRow < startRow)) {
		throw new InvalidArgumentError(`Range ends before it starts: [${startCol}, ${startRow}, ${endCol}, ${endRow}]`);
	}
	return [startCol, startRow, endCol, endRow];
}

export function parseRangeReferenceExact(range: RangeRef, worksheet: Worksheet): [number, number, number, number] {
	let [ac, ar, bc, br] = parseRangeReference(range);

	ac = ac ?? worksheet.cells.minDataColumn;
	ar = ar ?? worksheet.cells.minDataRow;
	bc = bc ?? worksheet.cells.maxDataColumn;
	br = br ?? worksheet.cells.maxDataRow;

	return [ac, ar, bc, br];
}

/**
 * Converts a column reference (e.g., "A", "Z") to its number (1-based).
 */
export function columnComponentToNumber(column: ColumnRef): ColumnNumber {
	let num = 0;
	for (let i = 0; i < column.length; i++) {
		num = num * 26 + (column.charCodeAt(i) - 65 + 1);
	}
	return num as ColumnNumber;
}

/**
 * Converts a row reference (string or number) to a number.
 */
export function rowComponentToNumber(row: RowRef): RowNumber {
	if (typeof row === "number") return row;
	const parsed = parseInt(row, 10);
	if (Number.isNaN(parsed)) throw new Error(`Invalid row component: ${row}`);
	return parsed as RowNumber;
}
