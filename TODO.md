```ts
readWorkbookByPath(parentRef, path): Promise<WorkbookId>
readWorkbook(driveItemRef): Promise<WorkbookId>
readFilteredWorkbook(parentRef, path, ?mapping?, ?filter?): Promise<WorkbookId> // Style not preserved

updateWorkbook(workbookId, workbook => {
    workbook.apply("sheet1!A1:B5", => {

    })
    const a = workbook.extract("sheet1:A1")
})

writeWorkbook(workbookId)
writeWorkbook(workbookId, parentRef, path, resolution): Promise<DriveItem & DriveItemRef>
``