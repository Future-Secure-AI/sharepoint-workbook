// async function createLocalWorkbook(scratchFolder: string, worksheets: Iterable<WriteWorksheet> | AsyncIterable<WriteWorksheet>, notifyProgress: (prepared: number | undefined, compressionRation: number | undefined, written: number | undefined, force: boolean) => void) {
// 	const rawFilePath = join(scratchFolder, `raw.xlsx`);
// 	const rawStream = createWriteStream(rawFilePath);
// 	const xls = new ExcelJS.stream.xlsx.WorkbookWriter({ stream: rawStream });
// 	let preparedCells = 0;
// 	for await (const { name, rows } of worksheets) {
// 		const worksheet = xls.addWorksheet(name);
// 		for await (const row of rows) {
// 			appendRow(worksheet, row);
// 			preparedCells += row.length;
// 			notifyProgress(preparedCells, undefined, undefined, false);
// 		}
// 		worksheet.commit();
// 	}
// 	await xls.commit();

// 	notifyProgress(undefined, undefined, undefined, true);
// 	return { localWorkbookPath: rawFilePath, preparedCells };
// }
