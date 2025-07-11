import type { WriteRow } from "./Row.ts";

export type Worksheet = {
	id: string;
	state: "visible" | "hidden" | "veryHidden";
} & WriteWorksheet;

export type WriteWorksheet = {
	name: string;
	rows: Iterable<WriteRow> | AsyncIterable<WriteRow>;
};
