Reset(ChkSelectAll);
Reset(TxtProject);
Reset(TxtPerson);
Reset(TxtFree1);
Reset(TxtFree2);
Reset(TxtSearch);
Reset(DpStartDate);
Reset(DpEndDate);
ForAll(
    gblExcelFiles As file,
    Patch(gblExcelFiles, file, {IsSelected: false})
);
Set(gblResult, "");
Set(LblStatus.Text, "リセットしました")
