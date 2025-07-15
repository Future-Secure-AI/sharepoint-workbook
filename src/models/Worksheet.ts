/**
 * Worksheet models.
 * @module Worksheet
 * @category Models
 */
import type { RowWrite } from "./Row.ts";

/**
 * Represents a worksheet in a workbook.
 * @property {string} id Unique identifier for the worksheet.
 * @property {"visible" | "hidden" | "veryHidden"} state Visibility state of the worksheet.
 * @property {string} name Name of the worksheet.
 * @property {Iterable<WriteRow> | AsyncIterable<WriteRow>} rows Rows in the worksheet.
 */
export type Worksheet = {
	id: string;
	state: "visible" | "hidden" | "veryHidden";
} & WorksheetWrite;

/**
 * Represents a worksheet to be written.
 * @property {string} name Name of the worksheet.
 * @property {Iterable<WriteRow> | AsyncIterable<WriteRow>} rows Rows to write.
 */
export type WorksheetWrite = {
	name: string;
	rows: Iterable<RowWrite> | AsyncIterable<RowWrite>;
};

export type WorksheetName = string & { readonly __brand: unique symbol };

export type WorksheetIndex = number;
