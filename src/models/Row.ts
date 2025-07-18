/**
 * Rows.
 * @module Row
 * @category Models
 */

/**
 * Represents a row index in a worksheet. This is a branded type to distinguish row indices from other numbers.
 * @typedef {number} RowIndex
 */
export type RowIndex = number & { readonly __brand: unique symbol };
