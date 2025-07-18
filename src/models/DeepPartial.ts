/**
 * Making all properties of a type optional, recursively.
 * @module DeepPartial
 * @category Models
 */

/**
 * Makes all properties of a type optional, recursively.
 * Useful for partial updates or patch operations on deeply nested objects.
 * @template T The type to make deeply partial.
 * @typedef {object} DeepPartial
 */
export type DeepPartial<T> = {
	[P in keyof T]?: T[P] extends object ? DeepPartial<T[P]> : T[P];
};
