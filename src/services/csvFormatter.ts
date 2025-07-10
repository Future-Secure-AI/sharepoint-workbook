import type { Row } from "@fast-csv/format";
import type { CsvFormatterStream } from "fast-csv";

export async function csvToBuffer(writer: CsvFormatterStream<Row, Row>): Promise<Buffer> {
    const chunks: Buffer[] = [];
    for await (const chunk of writer) {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
    }
    const buff = Buffer.concat(chunks);
    return buff;
}
