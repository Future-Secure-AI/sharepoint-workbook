import type { Cell } from "microsoft-graph/Cell";
import fs from "node:fs";
import https from "node:https";
import os from "node:os";
import path from "node:path";
import zlib from "node:zlib";

const LARGE_CSV_URL = "https://github.com/DataTalksClub/nyc-tlc-data/releases/download/yellow/yellow_tripdata_2019-03.csv.gz";
const LARGE_CSV_CACHE = path.join(os.tmpdir(), "730MB.csv.gz.tmp");

export async function getMemoryLimitMB(): Promise<number> {
    const v8 = await import("node:v8");
    return v8.getHeapStatistics().heap_size_limit / 1024 / 1024
}

export async function* getLargeSet(maxRows?: number): AsyncGenerator<Partial<Cell>[]> {
    // Download the file if it doesn't exist
    if (!fs.existsSync(LARGE_CSV_CACHE) || fs.statSync(LARGE_CSV_CACHE).size === 0) {
        console.log("Downloading large CSV dataset...");
        const file = fs.createWriteStream(LARGE_CSV_CACHE);
        await downloadWithRedirects(LARGE_CSV_URL, file);
        console.log("Download complete.");
    }

    const fileStream = fs.createReadStream(LARGE_CSV_CACHE);
    const gunzip = zlib.createGunzip();
    const decompressedStream = fileStream.pipe(gunzip);
    let leftover = "";
    let count = 0;
    for await (const chunk of decompressedStream) {
        const lines = (leftover + chunk.toString()).split("\n");
        leftover = lines.pop() ?? "";
        for (const line of lines) {
            if (!line.trim() || !line.includes(",")) continue;
            if (maxRows !== undefined && count++ >= maxRows) return;
            yield line.split(",").map((value) => ({ value }));
        }
    }
}

async function downloadWithRedirects(url: string, dest: fs.WriteStream, maxRedirects = 5): Promise<void> {
    return new Promise((resolve, reject) => {
        const options = {
            headers: {
                "User-Agent": "Node.js",
            },
        };
        const doRequest = (currentUrl: string, redirectsLeft: number) => {
            https
                .get(currentUrl, options, (response) => {
                    if (response.statusCode === 302 && response.headers.location && redirectsLeft > 0) {
                        // Follow redirect
                        doRequest(response.headers.location, redirectsLeft - 1);
                        response.resume();
                        return;
                    }
                    if (response.statusCode !== 200) {
                        reject(new Error(`Failed to download file: HTTP ${response.statusCode}`));
                        response.resume();
                        return;
                    }
                    response.pipe(dest);
                    dest.on("finish", () => {
                        dest.close((err) => {
                            if (err) reject(err);
                            else resolve();
                        });
                    });
                    dest.on("error", reject);
                })
                .on("error", reject);
        };
        doRequest(url, maxRedirects);
    });
}
