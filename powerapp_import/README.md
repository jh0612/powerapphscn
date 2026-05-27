# PowerApps 可导入资产（按手顺书实现）

本目录提供一套可直接用于 PowerApps/Power Automate/Office Scripts 的实现资产：

- `office-scripts/`：2 个 Office Scripts（TypeScript）
- `powerfx/`：各控件属性对应的 Power Fx 公式
- `sample-data/`：可直接上传到 SharePoint 的伪造 Excel 测试数据

## 1) Office Scripts 导入

在 Excel for Web →「自动化」→「新脚本」，分别粘贴：

- `office-scripts/CalculateWorkHours.ts`
- `office-scripts/WriteToTemplateWithMapping.ts`

脚本名分别保存为：

- `CalculateWorkHours`
- `WriteToTemplateWithMapping`

## 2) PowerApps 公式粘贴

按文件名对应控件属性粘贴（例如 `BtnExecute.OnSelect.fx` 粘贴到 `BtnExecute` 的 `OnSelect`）。

关键文件：

- `App.OnStart.fx`
- `DdLocation.OnChange.fx`
- `GalFiles.Items.fx`
- `BtnExecute.OnSelect.fx`
- `BtnClear.OnSelect.fx`
- `BtnOpenOutput.OnSelect.fx`
- `LblStatus.Text.fx`（先绑定到 `gblStatus` 变量）

> 注意：`BtnOpenOutput.OnSelect.fx` 中的 SharePoint URL 需要替换成你自己的站点地址。

## 3) 假数据 Excel 上传

请在 SharePoint 文档库 `勤怠データ` 中按下列结构上传：

- `sample-data/拠点別パラメータ.xlsx`
- `sample-data/単価表.xlsx`
- `sample-data/勤怠集計_テンプレート.xlsx`
- `sample-data/北京/北京_202605.xlsx`
- `sample-data/大連/大連_202605.xlsx`
- `sample-data/河南/河南_202605.xlsx`

## 4) Flow 对接约定

Flow 1：`GetExcelFilesByLocation`
- 输入：`LocationName`
- 输出：`files`

Flow 2：`BatchProcessAttendance`
- 输入顺序：`FileIds, LocationName, ProjectName, FreeInput1, FreeInput2`
- 输出：`resultFileUrl`

`BtnExecute` 中已按以上顺序调用：

`BatchProcessAttendance.Run(selectedIds, loc, prj, f1, f2)`

