/**
 * Find a worksheet in a workbook by name.
 * @module findWorksheet
 * @category Tasks
 */
import type { Worksheet } from "aspose.cells.node";
import NotFoundError from "microsoft-graph/NotFoundError";
import picomatch from "picomatch/lib/picomatch.js";
import type { Workbook } from "../models/Workbook.ts";

/**
 * Finds the first worksheet in the workbook whose name matches the given string or glob pattern (case-insensitive).
 * @param {Workbook} workbook The workbook to search.
 * @param {string} search The worksheet name or glob pattern to match.
 * @returns {Worksheet} The first matching worksheet.
 * @throws {NotFoundError} If no worksheet matches the pattern.
 */
export default function findWorksheet(workbook: Workbook, search: string): Worksheet {
	const matcher = picomatch(search, { nocase: true });

	for (let i = 0; i < workbook.worksheets.count; i++) {
		const worksheet = workbook.worksheets.get(i);
		if (matcher(worksheet.name)) {
			return worksheet;
		}
	}

	throw new NotFoundError(`Worksheet not found matching '${search}'`);
}
