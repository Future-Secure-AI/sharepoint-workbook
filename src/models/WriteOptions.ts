/**
 * Options for writing a workbook file.
 * @module WriteOptions
 * @category Models
 * @property {"fail" | "replace" | "rename"} [ifExists] Behavior if the file already exists.
 * @property {(bytes: number): void} [progress] Progress callback, receives bytes processed.
 * @property {number} [maxChunkSize] Maximum chunk size in bytes for writing.
 */
export type WriteOptions = {
	ifExists?: "fail" | "replace" | "rename";
	progress?: (bytes: number) => void;
	maxChunkSize?: number;
};
