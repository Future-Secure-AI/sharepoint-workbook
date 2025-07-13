export type WriteOptions = {
	ifExists?: "fail" | "replace" | "rename";
	progress?: (bytes: number) => void;
};
