import type { DriveItem } from "@microsoft/microsoft-graph-types";
import type { DriveItemRef } from "microsoft-graph/DriveItem";
import getDriveItem from "microsoft-graph/getDriveItem";
import InvalidArgumentError from "microsoft-graph/InvalidArgumentError";
import streamDriveItemContent from "microsoft-graph/streamDriveItemContent";
import { extname, join } from "node:path";
import type { OpenRef } from "../models/Open.ts";
import { createOpenId, getWorkbookFolder } from "../services/workingFolder.ts";

export default async function readWorkbook(itemRef: DriveItemRef & Partial<DriveItem>): Promise<OpenRef> {
	const id = createOpenId();
	let name = itemRef.name;
	if (!name) {
		const item = await getDriveItem(itemRef);
		name = item.name ?? "";
	}

	const extension = extname(name).toLowerCase();

	const folder = await getWorkbookFolder(id);
	const targetFileName = join(folder, "0");
	const stream = await streamDriveItemContent(itemRef);

	if (extension === ".xlsx") {
		// TODO: Write stream to targetFileName without buffering the entire file in memory
	} else if (extension === ".csv") {
		// TODO: Convert to XLSX and write to targetFileName without buffering the entire file in memory
	} else {
		throw new InvalidArgumentError(`Unsupported file extension "${extension}".`);
	}

	return {
		id,
		itemRef,
	};
}

// const xls = new ExcelJS.stream.xlsx.WorkbookReader(stream, {});
// const worksheets: AsyncGenerator<Worksheet> = (async function* () {
// 	for (const ws of xls.worksheets) {
// 		const rows: AsyncGenerator<Row> = (async function* () {
// 			for (let i = 1; i <= ws.rowCount; i++) {
// 				const r = ws.getRow(i);
// 				const values = Array.isArray(r.values) ? r.values.slice(1) : [];
// 				yield values.map(
// 					(v) =>
// 						({
// 							text: "",
// 							value: (v ?? "") as CellValue,
// 							format: "" as CellFormat,
// 							merge: {},
// 							alignment: {},
// 							borders: {},
// 							fill: {},
// 							font: {},
// 						}) satisfies Cell,
// 				);
// 			}
// 		})();

// 		yield {
// 			name: ws.name,
// 			rows: rows,
// 			id: ws.id?.toString?.() ?? ws.name,
// 			state: ws.state ?? "visible",
// 		} satisfies Worksheet;
// 	}
// })();
