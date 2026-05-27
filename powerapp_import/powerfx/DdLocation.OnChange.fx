UpdateContext({locLoading: true});
Reset(ChkSelectAll);
Set(
    gblExcelFiles,
    AddColumns(
        GetExcelFilesByLocation.Run(DdLocation.Selected.Value).files,
        "IsSelected",
        false
    )
);
UpdateContext({locLoading: false});
Set(
    gblStatus,
    DdLocation.Selected.Value & " のファイル一覧を取得しました（" & CountRows(gblExcelFiles) & "件）"
)
