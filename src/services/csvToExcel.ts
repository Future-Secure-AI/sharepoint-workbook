/**
 * Converting CSV to Excel format.
 * @module csvToExcel
 * @category Services
 */
import { parse } from "fast-csv";
import he from "he";
import { createWriteStream } from "node:fs";
import { Readable } from "node:stream";
import Yazl from "yazl";
import type { LocalFilePath } from "../models/LocalFilePath.ts";
import type { WorksheetName } from "../models/Worksheet.ts";

export type Options = {
	worksheetName?: WorksheetName;
	/**
	 * Compression level for the XLSX zip file (0-9). Default: 6
	 */
	compressionLevel?: number;
};

// TODO: the shared string lookup isn't bounded. So enough unique strings will cause it to run out of memory.

/**
 * Convert a CSV stream to an Excel file.
 * @remarks Why build this when ExcelJS supports stream writing? We'll, ExcelJS has a bug where it corrupts large files when stream writing. So here we go!
 * @param stream
 * @param outputFile
 * @param options
 */
export async function csvToExcel(stream: Readable, outputFile: LocalFilePath, options: Options = {}): Promise<void> {
	const { worksheetName = "Sheet1", compressionLevel = 6 } = options;
	const safeSheetName = he.encode(worksheetName, { useNamedReferences: true });
	const sheetFileName = "sheet1.xml";
	const maxSharedStringLength = 8; // Any longer strings are unlikely to be repeated

	const xmlTemplates = {
		contentTypes: () =>
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n  <Default Extension="xml" ContentType="application/xml"/>\n  <Override PartName="/xl/worksheets/${sheetFileName}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n</Types>`,
		rels: () =>
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>\n</Relationships>`,
		workbook: () =>
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n  <sheets>\n    <sheet name="${safeSheetName}" sheetId="1" r:id="rId1"/>\n  </sheets>\n</workbook>`,
		workbookRels: () =>
			`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/${sheetFileName}"/>\n</Relationships>`,
		sheetHeader: () => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${safeSheetName}"><sheetData>`,
		sheetFooter: () => `</sheetData></worksheet>`,
	};

	const sharedStringMap = new Map<string, number>();
	const sharedStrings: string[] = [];

	const zipfile = new Yazl.ZipFile();
	zipfile.addBuffer(Buffer.from(xmlTemplates.contentTypes(), "utf8"), "[Content_Types].xml", { compressionLevel });
	zipfile.addBuffer(Buffer.from(xmlTemplates.rels(), "utf8"), "_rels/.rels", { compressionLevel });
	zipfile.addBuffer(Buffer.from(xmlTemplates.workbook(), "utf8"), "xl/workbook.xml", { compressionLevel });
	zipfile.addBuffer(Buffer.from(xmlTemplates.workbookRels(), "utf8"), "xl/_rels/workbook.xml.rels", { compressionLevel });

	let rowIndex = 1;
	let pendingRows: string[] = [];
	const BATCH_SIZE = 1000; // Process rows in batches to control memory usage

	const worksheetStream = new Readable({
		read() {
			// This will be called when the stream wants more data
		},
	});

	worksheetStream.push(xmlTemplates.sheetHeader());

	const csvParsePromise = new Promise<void>((resolve, reject) => {
		stream
			.pipe(parse())
			.on("error", reject)
			.on("data", (cells: string[]) => {
				let rowXml = `<row r=\"${rowIndex}\">`;
				let hasValue = false;
				for (let c = 0; c < cells.length; ++c) {
					const col = columnName(c + 1);
					const cellValue = cells[c];
					if (!cellValue) continue;
					hasValue = true;
					const strValue = String(cellValue);
					if (strValue.length <= maxSharedStringLength) {
						let idx = sharedStringMap.get(strValue);
						if (idx === undefined) {
							idx = sharedStrings.length;
							sharedStringMap.set(strValue, idx);
							sharedStrings.push(strValue);
						}
						rowXml += `<c r=\"${col}${rowIndex}\" t=\"s\"><v>${idx}</v></c>`;
					} else {
						rowXml += `<c r=\"${col}${rowIndex}\" t=\"inlineStr\"><is><t>${he.encode(strValue, { useNamedReferences: true })}</t></is></c>`;
					}
				}
				rowXml += `</row>`;
				if (hasValue) {
					pendingRows.push(rowXml);

					if (pendingRows.length >= BATCH_SIZE) {
						worksheetStream.push(pendingRows.join(""));
						pendingRows = [];
					}
				}
				rowIndex++;
			})
			.on("end", () => {
				if (pendingRows.length > 0) {
					worksheetStream.push(pendingRows.join(""));
				}
				worksheetStream.push(xmlTemplates.sheetFooter());
				worksheetStream.push(null); // Signal end of stream
				resolve();
			});
	});

	zipfile.addReadStream(worksheetStream, `xl/worksheets/${sheetFileName}`, { compressionLevel });

	await csvParsePromise;

	const sharedStringsXml =
		`<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n` +
		`<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"${sharedStrings.length}\" uniqueCount=\"${sharedStrings.length}\">` +
		sharedStrings.map((s) => `<si><t>${he.encode(s, { useNamedReferences: true })}</t></si>`).join("") +
		`</sst>`;

	zipfile.addBuffer(Buffer.from(sharedStringsXml, "utf8"), "xl/sharedStrings.xml", { compressionLevel });

	await new Promise<void>((resolve, reject) => {
		zipfile.outputStream.pipe(createWriteStream(outputFile)).on("close", resolve).on("error", reject);
		zipfile.end();
	});
}
function columnName(n: number): string {
	// TODO: potentially duplicated. TODO: Unit test.
	let s = "";
	while (n > 0) {
		n--;
		s = String.fromCharCode(65 + (n % 26)) + s;
		n = Math.floor(n / 26);
	}
	return s;
}
