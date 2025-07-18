/**
 * Worksheets.
 * @module Worksheet
 * @category Models
 */

/**
 * Represents a worksheet name. This is a branded type to distinguish worksheet names from other strings.
 * @typedef {string} WorksheetName
 */
export type WorksheetName = string & { readonly __brand: unique symbol };
