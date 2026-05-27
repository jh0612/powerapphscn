SortByColumns(
    Filter(
        gblExcelFiles,
        IsBlank(TxtSearch.Text) || TxtSearch.Text in Name
    ),
    "Name",
    Ascending
)
