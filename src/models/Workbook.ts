import type { Worksheet, WriteWorksheet } from "./Worksheet.ts";

export type Workbook = {
	name: string;
	worksheets: Iterable<Worksheet> | AsyncIterable<Worksheet>;
};

export type WriteWorkbook = {
	name: string;
	worksheets: Iterable<WriteWorksheet> | AsyncIterable<WriteWorksheet>;
};
