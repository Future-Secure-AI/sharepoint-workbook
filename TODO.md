```ts
await readWorkbookByPathWithFiltering(parentRef, path, ?mapping?, ?filter?): Promise<Handle> // Style not preserved
```

```ts
await transactWorkbook(workbookId, workbook => {
    // TODO: Sane creation of cells with and without formatting
    workbook.listWorkbooks()

    workbook.insertCells(origin: CellRef, direction: InsertDirection, cells: CellValue | Partial<Cell> | (CellValue | Partial<Cell>)[][]) // <=== Use ExcelJS's `Cell` directly?
    workbook.updateCells(origin: CellRef, cells: CellValue | Partial<Cell> | (CellValue | Partial<Cell>)[][])
    workbook.deleteCells(ref: PartialRef, direction: DeletionDirection)
    const b = workbook.readCells(ref: PartialRef): Cell[][]
})
```

{
    sheet: "Sheet1"
    start: "A1
    end: "B3"
}
[sheet, start, end]
