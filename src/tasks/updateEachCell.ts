import type { Worksheet } from "aspose.cells.node";
import type { Cell, CellValue } from "../models/Cell.ts";
import type { DeepPartial } from "../models/DeepPartial.ts";
import type { RangeRef } from "../models/Reference.ts";
import { writeCell } from "../services/cell.ts";
import { parseRangeReferenceExact } from "../services/reference.ts";

export function updateEachCell(worksheet: Worksheet, range: RangeRef, write: CellValue | DeepPartial<Cell>): void {
	const [ac, ar, bc, br] = parseRangeReferenceExact(range, worksheet);

	for (let r = ar; r <= br; r++) {
		for (let c = ac; c <= bc; c++) {
			writeCell(worksheet, r, c, write);
		}
	}
}
