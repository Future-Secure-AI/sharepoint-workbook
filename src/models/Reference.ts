type Letter = "A" | "B" | "C" | "D" | "E" | "F" | "G" | "H" | "I" | "J" | "K" | "L" | "M" | "N" | "O" | "P" | "Q" | "R" | "S" | "T" | "U" | "V" | "W" | "X" | "Y" | "Z";

/**
 * Represents a column in a worksheet.
 * @remarks Only single letter columns are supported here due to TypeScript having a meltdown with longer column refs.
 */
export type ColumnComponent = `${Letter}`; // | `${Letter}${Letter}` | `${Letter}${Letter}${Letter}`;
export type RowComponent = `${number}` | number;
export type CellComponent = `${ColumnComponent}${RowComponent}`;

export type CellRef = `${string}!${CellComponent}` | [worksheet: string, cell: CellComponent];
export type RangeRef = `${string}!${ColumnComponent | RowComponent | CellComponent | ""}:${ColumnComponent | RowComponent | CellComponent | ""}` | [worksheetOrNamedRange: string, start: ColumnComponent | RowComponent | CellComponent | null, end: ColumnComponent | RowComponent | CellComponent | null];
export type ExplicitRangeRef = `${string}!${CellComponent}:${CellComponent}` | [worksheet: string, start: CellComponent, end: CellComponent];

export type ColumnNumber = number;
export type RowNumber = number;
