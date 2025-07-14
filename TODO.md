```ts
readWorkbookByPathWithFiltering(parentRef, path, ?mapping?, ?filter?): Promise<Handle> // Style not preserved
```

```ts
updateWorkbook(workbookId, workbook => {
    workbook.apply("sheet1!A1:B5", => {

    })
    const a = workbook.extract("sheet1:A1")
})
```