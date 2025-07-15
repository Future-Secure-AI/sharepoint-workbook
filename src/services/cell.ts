import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import type { CellValue, CellWrite } from "../models/Cell.ts";

export function normalizeCellWrite(cell: CellValue | CellWrite): CellWrite {
	const type = typeof cell;
	if (type === "string" || type === "number" || type === "boolean" || cell instanceof Date || cell === null) {
		return {
			value: cell as string | number | boolean | Date | null,
		};
	}
	if (type !== "object") {
		throw new InvalidArgumentError(`Unsupported cell type '${type}'.`);
	}

	return cell as CellWrite;
}
