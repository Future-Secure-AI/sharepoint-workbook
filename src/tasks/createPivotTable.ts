import type { Worksheet } from "aspose.cells.node";
import AsposeCells from "aspose.cells.node";
import type { CellRef, Ref } from "../models/Reference.ts";

export type CreatePivotTableOptions = {
	origin?: CellRef;
	name?: string;
};

export default function createPivotTable(worksheet: Worksheet, sourceWorksheet: Worksheet, sourceRange: Ref, options: CreatePivotTableOptions = {}): AsposeCells.PivotTable {
	const { origin = "A1", name = "PivotTable1" } = options;

	const source = `'${sourceWorksheet.name}'!${sourceRange}`;
	const pivotTables = worksheet.pivotTables;
	const pivotIndex = pivotTables.add(source, origin, name);
	const pivotTable = pivotTables.get(pivotIndex);

	pivotTable.autoFormatType = AsposeCells.PivotTableAutoFormatType.Report6; // TODO: Make option

	pivotTable.addFieldToArea(AsposeCells.PivotFieldType.Row, "VendorID");
	pivotTable.addFieldToArea(AsposeCells.PivotFieldType.Data, "fare_amount");

	return pivotTable;
}
