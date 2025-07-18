/**
 * References to one or more cells in a worksheet.
 * @module Reference
 * @category Models
 */

/**
 * Represents a column in a worksheet.
 * @remarks Only the first columns are covered by TypeScript type checking due to TypeScript complexity limitations, however all columns are supported at runtime.
 * @typedef {string} ColumnRef
 */
export type ColumnRef = `${Letter}`; // | `${Letter}${Letter}` | `${Letter}${Letter}${Letter}`;

/**
 * Represents a row in a worksheet. Can be a string or number.
 * @typedef {string|number} RowRef
 */
export type RowRef = `${number}` | number;

/**
 * Represents a cell reference in a worksheet (e.g., "A1").
 * @typedef {string} CellRef
 */
export type CellRef = `${ColumnRef}${RowRef}`;

/**
 * Represents a column, row, or cell reference.
 * @typedef {ColumnRef|RowRef|CellRef} ColumnRowOrCell
 * @internal
 */
export type ColumnOrRowOrCell = ColumnRef | RowRef | CellRef;

/**
 * Represents a worksheet reference, which can be a single column, row, cell, a range string, or a tuple.
 * @typedef {string|[ColumnOrRowOrCell|null, ColumnOrRowOrCell|null]} Ref
 */
export type Ref = ColumnOrRowOrCell | `${ColumnOrRowOrCell | ""}:${ColumnOrRowOrCell | ""}` | [start: ColumnOrRowOrCell | null, end: ColumnOrRowOrCell | null];

/**
 * Represents an explicit cell range reference (e.g., "A1:B2" or a tuple).
 * @typedef {string|[CellRef, CellRef]} ExplicitRef
 */
export type ExplicitRef = `${CellRef}:${CellRef}` | [start: CellRef, end: CellRef];

/**
 * Represents a single uppercase column letter (A-Z).
 * @typedef {('A'|'B'|'C'|'D'|'E'|'F'|'G'|'H'|'I'|'J'|'K'|'L'|'M'|'N'|'O'|'P'|'Q'|'R'|'S'|'T'|'U'|'V'|'W'|'X'|'Y'|'Z')} Letter
 */
export type Letter = "A" | "B" | "C" | "D" | "E" | "F" | "G" | "H" | "I" | "J" | "K" | "L" | "M" | "N" | "O" | "P" | "Q" | "R" | "S" | "T" | "U" | "V" | "W" | "X" | "Y" | "Z";
