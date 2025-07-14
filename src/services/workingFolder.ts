import { getEnvironmentVariable } from "microsoft-graph/dist/cjs/services/environmentVariable";
import { promises as fs } from "node:fs";
import { mkdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { HandleId } from "../models/Handle.ts";
import type { LocalFilePath } from "../models/LocalFilePath.ts";

async function getRootFolder(): Promise<string> {
	return getEnvironmentVariable("WORKING_FOLDER", join(tmpdir(), "sharepoint-workbook-working")) as LocalFilePath;
}

async function getWorkbookFolder(openId: HandleId): Promise<LocalFilePath> {
	const workingFolder = await getRootFolder();
	const scratchFolder = join(workingFolder, openId) as LocalFilePath;
	await mkdir(scratchFolder, { recursive: true });
	return scratchFolder;
}

export function createHandleId(): HandleId {
	return crypto.randomUUID() as HandleId;
}

export async function getLatestRevisionFilePath(id: HandleId): Promise<LocalFilePath> {
	const folder = await getWorkbookFolder(id);

	const files = await fs.readdir(folder);
	const numericFiles = files
		.filter((name) => /^\d+$/.test(name))
		.map(Number)
		.sort((a, b) => a - b);

	if (numericFiles.length === 0) {
		throw new Error("No numeric files found in folder");
	}

	const highestFile = String(numericFiles[numericFiles.length - 1]);
	const highestFilePath = join(folder, highestFile) as LocalFilePath;
	return highestFilePath;
}

export async function getNextRevisionFilePath(id: HandleId): Promise<LocalFilePath> {
	const folder = await getWorkbookFolder(id);
	const files = await fs.readdir(folder);
	const numericFiles = files.filter((name) => /^\d+$/.test(name)).map(Number);

	const nextNumber = numericFiles.length === 0 ? 0 : Math.max(...numericFiles) + 1;
	return join(folder, String(nextNumber)) as LocalFilePath;
}
