type Letter = "A" | "B" | "C" | "D" | "E" | "F" | "G" | "H" | "I" | "J" | "K" | "L" | "M" | "N" | "O" | "P" | "Q" | "R" | "S" | "T" | "U" | "V" | "W" | "X" | "Y" | "Z";

/**
 * Represents a column in a worksheet.
 * @remarks Only the first columns are covered by TypeScript type checking due to TypeScript complexity limitations, however all columns are supported at runtime.
 */
export type ColumnRef = `${Letter}`; // | `${Letter}${Letter}` | `${Letter}${Letter}${Letter}`;
export type RowRef = `${number}` | number;
export type CellRef = `${ColumnRef}${RowRef}`;
export type Ref = ColumnRef | RowRef | CellRef;

export type RangeRef = `${Ref | ""}:${Ref | ""}` | [start: Ref | null, end: Ref | null];
export type ExplicitRangeRef = `${CellRef}:${CellRef}` | [start: CellRef, end: CellRef];
