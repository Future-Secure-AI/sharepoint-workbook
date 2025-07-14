/**
 * Reference to an opened workbook.
 * @module Handle
 * @category Models
 */
import type { DriveItemRef } from "microsoft-graph/DriveItem";

/**
 * A reference to an opened workbook.
 * @property id Unique identifier for the handle.
 * @property itemRef (Optional) Reference to the associated DriveItem in Microsoft Graph.
 */
export type Handle = {
	id: HandleId;
	itemRef?: DriveItemRef;
};

/**
 * Unique Handle identifier.
 */
export type HandleId = string & {
	readonly __brand: unique symbol;
};
