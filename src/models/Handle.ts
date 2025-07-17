/**
 * Reference to an opened workbook.
 * @module Handle
 * @category Models
 */

import type AsposeCells from "aspose.cells.node";
import type { DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";

/**
 * A reference to an opened workbook.
 * @property localFilePath Unique identifier for the handle.
 * @property remoteItemRef (Optional) Reference to the associated DriveItem in Microsoft Graph.
 */
export type Handle = {
	workbook: AsposeCells.Workbook;
	remoteItemRef?: DriveItemRef;
};
