Set(
    gblExcelFiles,
    Table(
        {
            Name: "",
            Id: "",
            IsSelected: false,
            Modified: Now()
        }
    )
);
Set(gblResult, "");
Set(
    gblHistory,
    Table(
        {DateTime: "", Location: "", Project: "", Status: ""}
    )
);
If(
    !IsBlank(SavedHistory),
    Set(gblHistory, SavedHistory)
)
