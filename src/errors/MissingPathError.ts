/**
 * Error thrown when attempting to save a file when it hasn't been "saved as", so no path is known.
 * @module MissingPathError
 * @category Errors
 */
export default class MissingPathError extends Error {
	public constructor(message: string) {
		super(message);
		this.name = "MissingPathError";
	}
}
