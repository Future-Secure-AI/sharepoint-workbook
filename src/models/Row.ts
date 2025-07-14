/**
 * Row models
 * @module Row
 * @category Models
 */
import type { Cell } from "microsoft-graph/dist/cjs/models/Cell";

/**
 * Represents a row in a worksheet.
 * @typedef {Cell[]}
 */
export type Row = Cell[];

/**
 * Represents a row to be written to a worksheet (partial cells allowed).
 * @typedef {Partial<Cell>[]}
 */
export type WriteRow = Partial<Cell>[];
