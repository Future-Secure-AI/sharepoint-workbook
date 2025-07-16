/**
 * Converting CSV to Excel format.
 * @module csvToExcel
 * @category Services
 */
import { parse } from "fast-csv";
import { createWriteStream } from "node:fs";
import { type Readable, PassThrough } from "node:stream";
import Yazl from "yazl";
import type { LocalFilePath } from "../models/LocalFilePath.ts";
import type { WorksheetName } from "../models/Worksheet.ts";

export type Options = {
	worksheetName?: WorksheetName;
};

/**
 * Convert a CSV stream to an Excel file.
 * @remarks Why build this when ExcelJS supports stream writing? We'll, ExcelJS has a bug where it corrupts large files when stream writing. So here we go!
 * @param stream
 * @param outputFile
 * @param options
 */
export async function csvToExcel(stream: Readable, outputFile: LocalFilePath, options: Options = {}): Promise<void> {
	const { worksheetName = "Sheet1" } = options;
	const safeSheetName = escapeXml(worksheetName);
	const sheetFileName = `sheet1.xml`;
	const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/xl/worksheets/${sheetFileName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
	<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>`;

	const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

	const workbookXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
	<sheets>
		<sheet name="${safeSheetName}" sheetId="1" r:id="rId1"/>
	</sheets>
</workbook>`;

	const workbookRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${sheetFileName}"/>
</Relationships>`;

	const zipfile = new Yazl.ZipFile();
	zipfile.addBuffer(Buffer.from(contentTypes, "utf8"), "[Content_Types].xml");
	zipfile.addBuffer(Buffer.from(rels, "utf8"), "_rels/.rels");
	zipfile.addBuffer(Buffer.from(workbookXml, "utf8"), "xl/workbook.xml");
	zipfile.addBuffer(Buffer.from(workbookRels, "utf8"), "xl/_rels/workbook.xml.rels");

	const sheetStream = new PassThrough();
	zipfile.addReadStream(sheetStream, `xl/worksheets/${sheetFileName}`);

	sheetStream.write(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${safeSheetName}"><sheetData>`);

	let rowIndex = 1;
	await new Promise<void>((resolve, reject) => {
		stream
			.pipe(parse())
			.on("error", (err) => {
				sheetStream.end();
				reject(err);
			})
			.on("data", (cells: string[]) => {
				sheetStream.write(`<row r=\"${rowIndex}\">`);
				for (let c = 0; c < cells.length; ++c) {
					const col = columnName(c + 1);
					const cellValue = cells[c] ?? "";
					sheetStream.write(`<c r=\"${col}${rowIndex}\" t=\"inlineStr\"><is><t>${escapeXml(String(cellValue))}</t></is></c>`);
				}
				sheetStream.write(`</row>`);
				rowIndex++;
			})
			.on("end", () => {
				sheetStream.write(`</sheetData></worksheet>`);
				sheetStream.end();
				resolve();
			});
	});

	// Finalize zip
	await new Promise<void>((resolve, reject) => {
		zipfile.outputStream.pipe(createWriteStream(outputFile)).on("close", resolve).on("error", reject);
		zipfile.end();
	});
}
function columnName(n: number): string {
	let s = "";
	while (n > 0) {
		n--;
		s = String.fromCharCode(65 + (n % 26)) + s;
		n = Math.floor(n / 26);
	}
	return s;
}
function escapeXml(text: string): string {
	return text.replace(/[<>&"']/g, (c) => {
		switch (c) {
			case "<":
				return "&lt;";
			case ">":
				return "&gt;";
			case "&":
				return "&amp;";
			case '"':
				return "&quot;";
			case "'":
				return "&apos;";
		}
		return c;
	});
}
