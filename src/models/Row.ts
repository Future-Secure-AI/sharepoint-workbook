/**
 * Row models
 * @module Row
 * @category Models
 */

import type { Cell, CellValue, CellWrite } from "./Cell.ts";
/**
 * Represents a row in a worksheet.
 */
export type Row = Cell[];

/**
 * Represents a row to be written to a worksheet (partial cells allowed).
 */
export type RowWrite = (CellValue | CellWrite)[];
