/**
 * Worksheet models.
 * @module Worksheet
 * @category Models
 */
import type { WriteRow } from "./Row.ts";

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
} & WriteWorksheet;

/**
 * Represents a worksheet to be written.
 * @property {string} name Name of the worksheet.
 * @property {Iterable<WriteRow> | AsyncIterable<WriteRow>} rows Rows to write.
 */
export type WriteWorksheet = {
	name: string;
	rows: Iterable<WriteRow> | AsyncIterable<WriteRow>;
};
