/**
 * Columns.
 * @module Column
 * @category Models
 */

/**
 * Represents a column index in a worksheet. This is a branded type to distinguish column indices from other numbers.
 * @typedef {number} ColumnIndex
 */
export type ColumnIndex = number & { readonly __brand: unique symbol };

export type ColumnName = string & { readonly __brand: unique symbol };
