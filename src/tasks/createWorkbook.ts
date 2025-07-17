/**
 * Filter out unwanted rows and columns from a workbook.
 * @module createWorkbook
 * @category Tasks
 */

import AsposeCells from "aspose.cells.node";
import type { Handle } from "../models/Handle.ts";

/**
 * Create a new blank workbook.
 */
export function createWorkbook(): Handle {
	const workbook = new AsposeCells.Workbook();

	return {
		workbook,
	};
}
