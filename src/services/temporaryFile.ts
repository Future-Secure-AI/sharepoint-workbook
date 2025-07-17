import { getEnvironmentVariable } from "microsoft-graph/dist/cjs/services/environmentVariable";
import { mkdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { LocalFilePath } from "../models/LocalFilePath.ts";

export async function getTemporaryFilePath(extension: string = ".tmp"): Promise<LocalFilePath> {
	const rootFolder = getEnvironmentVariable("WORKING_FOLDER", join(tmpdir(), "sharepoint-workbook")) as string;
	await mkdir(rootFolder, { recursive: true });
	const fileName = `${crypto.randomUUID()}${extension}`;
	const tempFilePath = join(rootFolder, fileName) as LocalFilePath;
	return tempFilePath;
}
