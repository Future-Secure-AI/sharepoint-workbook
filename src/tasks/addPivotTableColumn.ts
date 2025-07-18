import AsposeCells from "aspose.cells.node";
import type { ColumnName } from "microsoft-graph/Column";

export default function addPivotTableColumn(pivotTable: AsposeCells.PivotTable, columnName: ColumnName): void {
	pivotTable.addFieldToArea(AsposeCells.PivotFieldType.Column, columnName);
}
