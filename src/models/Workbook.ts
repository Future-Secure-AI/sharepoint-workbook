/**
 * Reference to an opened workbook.
 * @module Handle
 * @category Models
 */

import type { DriveItem } from "@microsoft/microsoft-graph-types";
import type AsposeCells from "aspose.cells.node";
import type { DriveItemRef } from "microsoft-graph/dist/cjs/models/DriveItem";

/**
 * A reference to an opened workbook.
 * @property localFilePath Unique identifier for the handle.
 * @property remoteItemRef (Optional) Reference to the associated DriveItem in Microsoft Graph.
 */
export type Workbook = AsposeCells.Workbook & {
	remoteItem?: DriveItem & DriveItemRef;
};
