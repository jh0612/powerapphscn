ForAll(
    gblExcelFiles As file,
    Patch(gblExcelFiles, file, {IsSelected: true})
)
