/**
 * Optimize a opened workbook file by recompressing with a specified compression level.
 * @module optimizeWorkbook
 * @category Tasks
 */

import { createWriteStream } from "node:fs";
import { stat } from "node:fs/promises";
import yauzl, { type Entry, type ZipFile as YauzlZipFile } from "yauzl";
import yazl, { type ZipFile as YazlZipFile } from "yazl";
import type { Handle } from "../models/Handle.ts";
import type { OptimizeOptions } from "../models/OptimizeOptions.ts";
import { getLatestRevisionFilePath, getNextRevisionFilePath } from "../services/workingFolder.ts";

/**
 * Optimizes an opened workbook by recompressing it.
 * @param {Handle} handle Reference to the opened workbook.
 * @param {OptimizeOptions} options Options for optimization, including compression level.
 * @returns {Promise<number>} The ratio of the output file size to the input file size.
 * @throws {Error} If the optimization fails.
 */
export default async function optimizeWorkbook(handle: Handle, options: OptimizeOptions = {}): Promise<number> {
	const { compressionLevel = 6 } = options;
	const latestFile = await getLatestRevisionFilePath(handle.id);
	const nextFile = await getNextRevisionFilePath(handle.id);

	await new Promise<void>((resolve, reject) => {
		yauzl.open(latestFile, { lazyEntries: true, autoClose: true }, (err: Error | null, zipfile: YauzlZipFile | undefined) => {
			if (err || !zipfile) return reject(err);
			const zipWriter: YazlZipFile = new yazl.ZipFile();
			const stream = createWriteStream(nextFile);
			zipWriter.outputStream.pipe(stream).on("close", resolve).on("error", reject);
			zipfile.readEntry();
			zipfile.on("entry", (entry: Entry) => {
				zipfile.openReadStream(entry, (err: Error | null, readStream: NodeJS.ReadableStream | undefined) => {
					if (err || !readStream) return reject(err);
					zipWriter.addReadStream(readStream, entry.fileName, {
						compress: true,
						compressionLevel,
						mtime: entry.getLastModDate ? entry.getLastModDate() : new Date(),
						mode: entry.externalFileAttributes >>> 16,
					});
					zipfile.readEntry();
				});
			});
			zipfile.on("end", () => {
				zipWriter.end();
			});
			zipfile.on("error", reject);
		});
	});

	const { size: inputSize } = await stat(latestFile);
	const { size: outputSize } = await stat(nextFile);
	return outputSize / inputSize;
}
