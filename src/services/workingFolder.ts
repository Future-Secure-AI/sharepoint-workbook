import { getEnvironmentVariable } from "microsoft-graph/environmentVariable";
import { promises as fs } from "node:fs";
import { mkdir } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import type { OpenId } from "../models/Open.ts";

async function getWorkingFolder(): Promise<string> {
	return getEnvironmentVariable("WORKING_FOLDER", join(tmpdir(), "sharepoint-workbook-working")) as string;
}

export async function getWorkbookFolder(openId: OpenId): Promise<string> {
	const workingFolder = await getWorkingFolder();
	const scratchFolder = join(workingFolder, openId);
	await mkdir(scratchFolder, { recursive: true });
	return scratchFolder;
}

export function createOpenId(): OpenId {
	return crypto.randomUUID() as OpenId;
}

export async function getLatest(id: OpenId) {
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
	const highestFilePath = join(folder, highestFile);
	return highestFilePath;
}
