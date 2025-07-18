/**
 * Utilities for parsing and resolving cell, row, column, and range references in worksheets.
 * @module reference
 * @category Services
 */
/** biome-ignore-all lint/complexity/useLiteralKeys: Impossible to avoid with RegEx */

import type { Worksheet } from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { ColumnIndex } from "../models/Column.ts";
import type { CellRef, ColumnRef, Ref, RowRef } from "../models/Reference.ts";
import type { RowIndex } from "../models/Row.ts";

// Matches cell refs like A1, Z99, etc.
const cellPattern = /^(?<col>[A-Z]{1,3})(?<row>\d{1,7})$/;
// Matches range refs like A1:C3, 1:5, A:C, etc.
const rangePattern = /^(?<startCol>[A-Z]+)?(?<startRow>\d+)?:(?<endCol>[A-Z]+)?(?<endRow>\d+)?$/;

/**
 * Parses a cell reference (e.g., "A1") into a tuple of 0-based column and row indices.
 * @param {CellRef} cell Reference string such as "B2".
 * @returns {[ColumnIndex, RowIndex]} 0-based column and row indices.
 * @throws {Error} If the cell reference format is invalid.
 */
export function parseCellRef(cell: CellRef): [ColumnIndex, RowIndex] {
	const match = cell.toString().match(cellPattern)?.groups;
	if (!match) throw new Error(`Invalid cell reference format: '${cell}'`);
	return [resolveColumnIndex(match["col"] as ColumnRef), resolveRowIndex(match["row"] as RowRef)];
}

/**
 * Converts a range reference (string or array) to a tuple of 0-based indices: [startCol, startRow, endCol, endRow].
 * Accepts cell, row, column, or range references in string or array form.
 * @param {Ref} range Reference such as "A1:C3", ["A", "C"], [1, 5], etc.
 * @returns {[number | null, number | null, number | null, number | null]} Resolved indices (null if not specified).
 * @throws {InvalidArgumentError|Error} If the reference is invalid or the range ends before it starts.
 */
export function parseRef(range: Ref): [ColumnIndex | null, RowIndex | null, ColumnIndex | null, RowIndex | null] {
	if (!Array.isArray(range)) {
		// Handle single column reference (e.g., "C")
		if (typeof range === "string" && /^[A-Z]+$/.test(range)) {
			const col = resolveColumnIndex(range as ColumnRef);
			return [col, null, col, null];
		}
		// Handle single row reference (e.g., 3 or "3")
		if ((typeof range === "number" && Number.isInteger(range)) || (typeof range === "string" && /^\d+$/.test(range))) {
			const row = resolveRowIndex(range as RowRef);
			return [null, row, null, row];
		}
		// Try cell reference first
		if (typeof range === "string") {
			const cellMatch = range.match(cellPattern)?.groups;
			if (cellMatch) {
				const col = resolveColumnIndex(cellMatch["col"] as ColumnRef);
				const row = resolveRowIndex(cellMatch["row"] as RowRef);
				return [col, row, col, row];
			}
			// Try range reference
			const match = range.match(rangePattern)?.groups;
			if (!match) throw new Error(`Invalid range reference format: ${range}`);
			const startCol = match["startCol"] ? resolveColumnIndex(match["startCol"] as ColumnRef) : null;
			const startRow = match["startRow"] ? resolveRowIndex(match["startRow"] as RowRef) : null;
			const endCol = match["endCol"] ? resolveColumnIndex(match["endCol"] as ColumnRef) : null;
			const endRow = match["endRow"] ? resolveRowIndex(match["endRow"] as RowRef) : null;
			if ((startCol !== null && endCol !== null && endCol < startCol) || (startRow !== null && endRow !== null && endRow < startRow)) {
				throw new InvalidArgumentError(`Range ends before it starts: [${startCol}, ${startRow}, ${endCol}, ${endRow}]`);
			}
			return [startCol, startRow, endCol, endRow];
		}
		throw new Error(`Invalid range reference: ${range}`);
	}

	// Now we know range is an array
	if (range.length !== 2) throw new Error(`Invalid range reference array: ${range}`);
	const [start, end] = range;

	// Row-only range: [number, number] or [string, string] with both numeric
	const isRowOnly = (v: unknown) => typeof v === "number" || (typeof v === "string" && /^\d+$/.test(v));
	// Column-only range: [string, string] with both uppercase letters
	const isColOnly = (v: unknown) => typeof v === "string" && /^[A-Z]+$/.test(v);

	if (isRowOnly(start) && isRowOnly(end)) {
		const startRow = resolveRowIndex(start as RowRef);
		const endRow = resolveRowIndex(end as RowRef);
		if (endRow < startRow) throw new InvalidArgumentError(`Range ends before it starts: [null, ${startRow}, null, ${endRow}]`);
		return [null, startRow, null, endRow];
	}
	if (isColOnly(start) && isColOnly(end)) {
		const startCol = resolveColumnIndex(start as ColumnRef);
		const endCol = resolveColumnIndex(end as ColumnRef);
		if (endCol < startCol) throw new InvalidArgumentError(`Range ends before it starts: [${startCol}, null, ${endCol}, null]`);
		return [startCol, null, endCol, null];
	}
	// Column-row range: [col, row] (e.g., ["A", 3] or ["A", "3"])
	if (isColOnly(start) && isRowOnly(end)) {
		const startCol = resolveColumnIndex(start as ColumnRef);
		const endRow = resolveRowIndex(end as RowRef);
		return [startCol, null, null, endRow];
	}
	// Row-column range: [row, col] (e.g., [1, "C"] or ["1", "C"])
	if (isRowOnly(start) && isColOnly(end)) {
		const startRow = resolveRowIndex(start as RowRef);
		const endCol = resolveColumnIndex(end as ColumnRef);
		return [null, startRow, endCol, null];
	}

	// Otherwise, treat as cell references
	const parseCell = (ref: Ref | null) => {
		if (ref == null) return [null, null];
		try {
			return parseCellRef(ref as CellRef);
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

/**
 * Resolves a range reference to concrete 0-based indices, filling in worksheet bounds for omitted values.
 * @param {Ref} range Reference to resolve.
 * @param {Worksheet} worksheet Worksheet to use for bounds.
 * @returns {[number, number, number, number]} Fully resolved indices: [startCol, startRow, endCol, endRow].
 */
export function parseRefResolved(range: Ref, worksheet: Worksheet): [ColumnIndex, RowIndex, ColumnIndex, RowIndex] {
	let [ac, ar, bc, br] = parseRef(range);

	ac = ac ?? (worksheet.cells.minDataColumn as ColumnIndex);
	ar = ar ?? (worksheet.cells.minDataRow as RowIndex);
	bc = bc ?? (worksheet.cells.maxDataColumn as ColumnIndex);
	br = br ?? (worksheet.cells.maxDataRow as RowIndex);

	return [ac, ar, bc, br];
}

/**
 * Converts a column reference (e.g., "A", "Z", "AA") to its 0-based column index.
 * @param {ColumnRef} column Column reference string.
 * @returns {ColumnIndex} 0-based column index.
 */
export function resolveColumnIndex(column: ColumnRef): ColumnIndex {
	let num = 0;
	for (let i = 0; i < column.length; i++) {
		num = num * 26 + (column.charCodeAt(i) - 65 + 1);
	}
	return (num - 1) as ColumnIndex;
}

/**
 * Converts a row reference (string or number) to a 0-based row index.
 * @param {RowRef} row Row reference (number or string).
 * @returns {RowIndex} 0-based row index.
 * @throws {Error} If the row reference is invalid.
 */
export function resolveRowIndex(row: RowRef): RowIndex {
	const parsed = Number.parseInt(row, 10);
	if (Number.isNaN(parsed)) throw new Error(`Invalid row component: ${row}`);
	return (parsed - 1) as RowIndex;
}
