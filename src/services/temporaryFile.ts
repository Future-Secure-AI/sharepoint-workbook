/**
 * Utilities for creating temporary file paths for workbook operations.
 * @module temporaryFile
 * @category Services
 */
import { getEnvironmentVariable } from "microsoft-graph/dist/cjs/services/environmentVariable";
import { mkdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { LocalFilePath } from "../models/LocalFilePath.ts";

/**
 * Generate a unique temporary file path, ensuring the directory exists.
 * @remarks The file is not created, only the path is generated. The root folder is determined by the `WORKING_FOLDER` environment variable if set, otherwise defaults to a subdirectory in the OS temp directory.
 * @param {string} [extension=".tmp"] - The file extension to use for the temporary file.
 * @returns {Promise<LocalFilePath>} The generated temporary file path.
 */
export async function getTemporaryFilePath(extension: string = ".tmp"): Promise<LocalFilePath> {
	const rootFolder = getEnvironmentVariable("WORKING_FOLDER", join(tmpdir(), "sharepoint-workbook")) as string;
	await mkdir(rootFolder, { recursive: true });
	const fileName = `${crypto.randomUUID()}${extension}`;
	const tempFilePath = join(rootFolder, fileName) as LocalFilePath;
	return tempFilePath;
}
