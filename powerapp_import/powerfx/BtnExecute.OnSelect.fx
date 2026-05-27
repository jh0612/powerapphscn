UpdateContext({locProcessing: true});
Set(gblStatus, "処理中...");

With(
    {
        selectedIds: Concat(
            Filter(gblExcelFiles, IsSelected),
            Id,
            ","
        ),
        loc: DdLocation.Selected.Value,
        prj: TxtProject.Text,
        f1: TxtFree1.Text,
        f2: TxtFree2.Text
    },
    If(
        IsBlank(selectedIds),
        Notify("ファイルを1件以上選択してください", NotificationType.Warning);
        Set(gblStatus, "ファイルが選択されていません"),
        If(
            Confirm(
                "【実行確認】" & Char(10) &
                "拠点：" & loc & Char(10) &
                "案件：" & prj & Char(10) &
                "ファイル数：" & CountRows(Filter(gblExcelFiles, IsSelected)) & "件" & Char(10) &
                Char(10) &
                "※入力ファイルは1ヶ月分の勤怠データとして処理されます。" & Char(10) &
                "実行しますか？"
            ),
            Set(
                gblResult,
                BatchProcessAttendance.Run(
                    selectedIds,
                    loc,
                    prj,
                    f1,
                    f2
                )
            );
            Notify("集計が完了しました", NotificationType.Success);
            Set(gblStatus, "処理完了：" & Text(Now(), "yyyy/MM/dd HH:mm"));
            Collect(
                gblHistory,
                {
                    DateTime: Text(Now(), "yyyy/MM/dd HH:mm"),
                    Location: loc,
                    Project: prj,
                    Status: "成功"
                }
            );
            SaveData(gblHistory, "SavedHistory"),
            Set(gblStatus, "キャンセルされました")
        )
    )
);

UpdateContext({locProcessing: false})
