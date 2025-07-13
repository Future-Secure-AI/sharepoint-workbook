// import type { DriveItem } from "@microsoft/microsoft-graph-types";
// import ExcelJS from "exceljs";
// import type { Cell, CellFormat, CellValue } from "microsoft-graph/Cell";
// import type { DriveRef } from "microsoft-graph/Drive";
// import type { DriveItemPath, DriveItemRef } from "microsoft-graph/DriveItem";
// import getDriveItemByPath from "microsoft-graph/getDriveItemByPath";
// import streamDriveItemContent from "microsoft-graph/streamDriveItemContent";
// import { basename } from "node:path";
// import type { Row } from "../models/Row.ts";
// import type { Workbook } from "../models/Workbook.js";
// import type { Worksheet } from "../models/Worksheet.ts";

// export default async function readWorkbook(parentRef: DriveRef | DriveItemRef, itemPath: DriveItemPath): Promise<Workbook & DriveItem & DriveItemRef> {
// 	const item = await getDriveItemByPath(parentRef, itemPath);
// 	const stream = await streamDriveItemContent(item);

// 	const xls = new ExcelJS.stream.xlsx.WorkbookReader(stream, {});

// 	const worksheets: AsyncGenerator<Worksheet> = (async function* () {
// 		for (const ws of xls.worksheets) {
// 			const rows: AsyncGenerator<Row> = (async function* () {
// 				for (let i = 1; i <= ws.rowCount; i++) {
// 					const r = ws.getRow(i);
// 					const values = Array.isArray(r.values) ? r.values.slice(1) : [];
// 					yield values.map(
// 						(v) =>
// 							({
// 								text: "",
// 								value: (v ?? "") as CellValue,
// 								format: "" as CellFormat,
// 								merge: {},
// 								alignment: {},
// 								borders: {},
// 								fill: {},
// 								font: {},
// 							}) satisfies Cell,
// 					);
// 				}
// 			})();

// 			yield {
// 				name: ws.name,
// 				rows: rows,
// 				id: ws.id?.toString?.() ?? ws.name,
// 				state: ws.state ?? "visible",
// 			} satisfies Worksheet;
// 		}
// 	})();

// 	const name = item.name ?? basename(itemPath);

// 	const workbook: Workbook = {
// 		name,
// 		worksheets,
// 	};

// 	return {
// 		...item,
// 		...workbook,
// 	};
// }
