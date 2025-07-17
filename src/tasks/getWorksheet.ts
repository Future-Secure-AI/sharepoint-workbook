import type AsposeCells from "aspose.cells.node";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { Workbook } from "../models/Workbook.ts";
import type { WorksheetName } from "../models/Worksheet.ts";

export default function getWorksheet(workbook: Workbook, name: WorksheetName): AsposeCells.Worksheet {
	const worksheet = workbook.worksheets.get(name);
	if (!worksheet) throw new InvalidArgumentError(`Worksheet not found: ${name}`);
	return worksheet;
}
