/**
 * Create a new worksheet in the given workbook with the specified name.
 * @module createWorksheet
 * @category Tasks
 */
import type { Workbook, Worksheet } from "aspose.cells.node";

/**
 * Adds a new worksheet to the workbook with the given name.
 * @param workbook The workbook to add the worksheet to.
 * @param name The name of the new worksheet.
 * @returns The newly created worksheet.
 */
export default function createWorksheet(workbook: Workbook, name: string): Worksheet {
	return workbook.worksheets.add(name);
}
