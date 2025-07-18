/**
 * Get a worksheet from a workbook by its exact name.
 * @module getWorksheet
 * @category Tasks
 */
import type AsposeCells from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { Workbook } from "../models/Workbook.ts";
import type { WorksheetName } from "../models/Worksheet.ts";

/**
 * Returns the worksheet with the given name from the workbook.
 * @param {Workbook} workbook Workbook to search.
 * @param {WorksheetName} name Name of the worksheet to retrieve (case insensitive).
 * @returns {AsposeCells.Worksheet} The worksheet with the specified name.
 * @throws {InvalidArgumentError} If the worksheet is not found.
 */
export default function getWorksheet(workbook: Workbook, name: WorksheetName): AsposeCells.Worksheet {
	const worksheet = workbook.worksheets.get(name);
	if (!worksheet) throw new InvalidArgumentError(`Worksheet not found: ${name}`);
	return worksheet;
}
