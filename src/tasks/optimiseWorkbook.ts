// import yauzl, { type Entry, type ZipFile as YauzlZipFile } from "yauzl";
// import yazl, { type ZipFile as YazlZipFile } from "yazl";
// async function optimiseWorkbook(inputFilePath: string, scratchFolder: string, compressionLevel: number, notifyProgress: (prepared: number | undefined, compressionRation: number | undefined, written: number | undefined, force: boolean) => void): Promise<string> {
// 	if (compressionLevel === 0) {
// 		notifyProgress(undefined, 1, undefined, true);
// 		return inputFilePath;
// 	}

// 	const outputFilePath = join(scratchFolder, `compressed.xlsx`);
// 	await new Promise<void>((resolve, reject) => {
// 		yauzl.open(inputFilePath, { lazyEntries: true, autoClose: true }, (err: Error | null, zipfile: YauzlZipFile | undefined) => {
// 			if (err || !zipfile) return reject(err);
// 			const zipWriter: YazlZipFile = new yazl.ZipFile();
// 			const stream = createWriteStream(outputFilePath);
// 			zipWriter.outputStream.pipe(stream).on("close", resolve).on("error", reject);
// 			zipfile.readEntry();
// 			zipfile.on("entry", (entry: Entry) => {
// 				zipfile.openReadStream(entry, (err: Error | null, readStream: NodeJS.ReadableStream | undefined) => {
// 					if (err || !readStream) return reject(err);
// 					zipWriter.addReadStream(readStream, entry.fileName, {
// 						compress: true,
// 						compressionLevel,
// 						mtime: entry.getLastModDate ? entry.getLastModDate() : new Date(),
// 						mode: entry.externalFileAttributes >>> 16,
// 					});
// 					zipfile.readEntry();
// 				});
// 			});
// 			zipfile.on("end", () => {
// 				zipWriter.end();
// 			});
// 			zipfile.on("error", reject);
// 		});
// 	});

// 	const { size: inputSize } = await fs.stat(inputFilePath);
// 	const { size: outputSize } = await fs.stat(outputFilePath);
// 	notifyProgress(undefined, outputSize / inputSize, undefined, true);
// 	return outputFilePath;
// }
