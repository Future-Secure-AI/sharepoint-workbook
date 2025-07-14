import type { WriteRow } from "../models/Row.ts";

export function asRows(rows: unknown[][]): AsyncIterable<WriteRow> {
	return (async function* () {
		for (const r of rows) yield r.map((cell) => ({ value: cell })) as WriteRow;
	})();
}
