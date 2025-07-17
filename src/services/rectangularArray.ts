import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";

export function ensureRectangularArray(array: unknown[][]): void {
	const rowCount = array.length;
	const colCount = array[0]?.length || 0;

	for (let i = 0; i < rowCount; i++) {
		if ((array[i]?.length || 0) !== colCount) {
			throw new InvalidArgumentError(`All rows in must have the same length. Row 0 has length ${colCount}, but row ${i} has length ${array[i]?.length || 0}.`);
		}
	}
}
