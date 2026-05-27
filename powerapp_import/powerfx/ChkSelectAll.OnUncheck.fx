ForAll(
    gblExcelFiles As file,
    Patch(gblExcelFiles, file, {IsSelected: false})
)
