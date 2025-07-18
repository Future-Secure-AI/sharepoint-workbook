/**
 * Streaming parameter utilities for workbook operations.
 * @module streamParameters
 * @category Services
 */

/**
 * Maximum amount of data (in bytes) to buffer in memory for streaming operations.
 * Used as the highWaterMark for streams to control memory usage.
 */
export const streamHighWaterMark = 1024 * 1024;
