import { getEnvironmentVariable } from "microsoft-graph/environmentVariable";
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
