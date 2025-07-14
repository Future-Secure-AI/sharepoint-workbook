/**
 * Utilities for manipulating rows.
 * @module Rows
 * @category Services
 */

import type { CellValue } from "microsoft-graph/dist/cjs/models/Cell";
import type { WriteRow } from "../models/Row.ts";

/**
 * Converts an array of arrays into an async iterable of WriteRow.
 */
export function arrayToRows(rows: CellValue[][]): AsyncIterable<WriteRow> {
	return (async function* () {
		for (const r of rows) yield r.map((cell) => ({ value: cell })) as WriteRow;
	})();
}
