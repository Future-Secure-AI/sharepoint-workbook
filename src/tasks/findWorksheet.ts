import type { Worksheet } from "aspose.cells.node";
import NotFoundError from "microsoft-graph/NotFoundError";
import picomatch from "picomatch/lib/picomatch.js";
import type { Workbook } from "../models/Workbook.ts";

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
