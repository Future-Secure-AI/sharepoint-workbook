/**
 * Workbook models.
 * @module Workbook
 * @category Models
 */
import type { Worksheet, WorksheetWrite } from "./Worksheet.ts";

/**
 * Represents a workbook with worksheets.
 * @property {string} name Name of the workbook.
 * @property {Iterable<Worksheet> | AsyncIterable<Worksheet>} worksheets Worksheets in the workbook.
 */
export type Workbook = {
	name: string;
	worksheets: Iterable<Worksheet> | AsyncIterable<Worksheet>;
};

/**
 * Represents a workbook to be written, with writeable worksheets.
 * @property {string} name Name of the workbook.
 * @property {Iterable<WriteWorksheet> | AsyncIterable<WriteWorksheet>} worksheets Worksheets to write.
 */
export type WorkbookWrite = {
	name: string;
	worksheets: Iterable<WorksheetWrite> | AsyncIterable<WorksheetWrite>;
};
