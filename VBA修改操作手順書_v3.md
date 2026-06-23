# VBA 勤怠集計ツール v3 修改操作手順書

> **目的**：将 `toolV2.xlsm` 升级为 `toolV2_v3.xlsm`，修复以下三个问题：
> 1. 选择据点后文件列表报错
> 2. 打开文件时先显示欢迎界面
> 3. 业务期间自动填写（日期选择辅助）
>
> **重要**：不要修改原 `toolV2.xlsm`，请先复制一份命名为 `toolV2_v3.xlsm`，在副本上操作。

---

## 准备工作：复制文件

1. 在文件资源管理器中找到 `toolV2.xlsm`
2. 右键 → 复制 → 粘贴，得到 `toolV2 - 副本.xlsm`
3. 将其重命名为 `toolV2_v3.xlsm`
4. 打开 `toolV2_v3.xlsm`，按 **Alt+F11** 打开 VBA 编辑器

---

## 修改 1：修复据点选择后文件列表报错

### 问题原因

`cboLocation_Change()` 事件切换据点时，虽然调用了 `RefreshFileList`，但**没有自动更新 `txtFolderPath.Text`**。  
`RefreshFileList` 函数只读取 `txtFolderPath.Text`，因此仍使用旧路径（或空路径），导致找不到文件或报错。

### 操作步骤

在 VBA 编辑器左侧项目树中，双击 **frmMain**，找到以下函数：

**找到这段原始代码（约第 83–90 行）：**

```vba
Private Sub cboLocation_Change()
    If cboLocation.ListIndex >= 0 Then
        SaveSetting "LastLocation", cboLocation.Text
        ' 拠点が変わったら担当者一覧を更新
        Call RefreshPersonList
        Call RefreshFileList
    End If
End Sub
```

**全部选中并替换为以下代码：**

```vba
Private Sub cboLocation_Change()
    On Error GoTo ErrHandler
    
    If cboLocation.ListIndex < 0 Then Exit Sub
    
    ' 保存最后选择的据点
    SaveSetting "LastLocation", cboLocation.Text
    
    ' 自动设置该据点对应的文件夹路径
    Dim autoPath As String
    autoPath = ThisWorkbook.Path & "\" & cboLocation.Text
    
    ' 仅当该文件夹存在时自动填入路径（否则保留手动路径）
    If Len(Dir(autoPath, vbDirectory)) > 0 Then
        txtFolderPath.Text = autoPath
        SaveSetting "LastFolderPath", autoPath
    End If
    
    ' 刷新担当者和文件列表
    Call RefreshPersonList
    Call SafeRefreshFileList
    
    Exit Sub

ErrHandler:
    ' 出错时静默清空列表，不弹窗
    lstFiles.Clear
    lblCount.Caption = "選択件数：0 件"
End Sub
```

---

### 同时替换 RefreshFileList 函数，增加容错处理

**找到这段原始代码（约第 344–415 行）：**

```vba
Private Sub RefreshFileList()
    lstFiles.Clear
    
    Dim folderPath As String
    folderPath = Trim(txtFolderPath.Text)
    
    If Len(folderPath) = 0 Then Exit Sub
    
    Dim files As Variant
    files = GetExcelFiles(folderPath)
    ...（后续代码省略）
End Sub
```

在该函数**正下方**（`End Sub` 之后）**新增**以下函数：

```vba
'--------------------------------------------------
' 安全版文件列表刷新（带完整错误处理）
'--------------------------------------------------
Private Sub SafeRefreshFileList()
    On Error GoTo ErrHandler
    
    lstFiles.Clear
    
    Dim folderPath As String
    folderPath = Trim(txtFolderPath.Text)
    
    ' 路径为空时，显示提示并退出
    If Len(folderPath) = 0 Then
        lblCount.Caption = "選択件数：0 件"
        Exit Sub
    End If
    
    ' 末尾の \ を除去して統一
    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If
    
    ' 文件夹不存在时，显示友好提示
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        lblCount.Caption = "選択件数：0 件（フォルダが存在しません）"
        Exit Sub
    End If
    
    ' 扫描 .xlsx 文件
    Dim fileName As String
    Dim fileCount As Long
    fileCount = 0
    
    fileName = Dir(folderPath & "\*.xlsx")
    Do While Len(fileName) > 0
        ' 跳过工具文件本身
        If LCase(fileName) <> LCase(ThisWorkbook.Name) Then
            Dim filePath As String
            filePath = folderPath & "\" & fileName
            
            On Error Resume Next
            Dim fso As Object
            Set fso = CreateObject("Scripting.FileSystemObject")
            Dim fileObj As Object
            Set fileObj = fso.GetFile(filePath)
            
            Dim fileSize As String
            If fileObj.Size < 1048576 Then
                fileSize = Format(fileObj.Size / 1024, "0") & " KB"
            Else
                fileSize = Format(fileObj.Size / 1048576, "0.0") & " MB"
            End If
            
            Dim modDate As String
            modDate = Format(fileObj.DateLastModified, "yyyy/mm/dd HH:MM")
            On Error GoTo ErrHandler
            
            lstFiles.AddItem fileName
            Dim rowIdx As Long
            rowIdx = lstFiles.ListCount - 1
            lstFiles.List(rowIdx, 1) = fileSize
            lstFiles.List(rowIdx, 2) = modDate
            fileCount = fileCount + 1
        End If
        fileName = Dir()
    Loop
    
    ' 扫描 .xls 文件（不含 .xlsm 等）
    fileName = Dir(folderPath & "\*.xls")
    Do While Len(fileName) > 0
        ' 跳过 .xlsm / .xlsb 等（Dir *.xls 会匹配所有以xls开头的扩展名）
        Dim ext As String
        ext = LCase(Right(fileName, 4))
        If ext = ".xls" Then
            If LCase(fileName) <> LCase(ThisWorkbook.Name) Then
                filePath = folderPath & "\" & fileName
                
                On Error Resume Next
                Set fso = CreateObject("Scripting.FileSystemObject")
                Set fileObj = fso.GetFile(filePath)
                If fileObj.Size < 1048576 Then
                    fileSize = Format(fileObj.Size / 1024, "0") & " KB"
                Else
                    fileSize = Format(fileObj.Size / 1048576, "0.0") & " MB"
                End If
                modDate = Format(fileObj.DateLastModified, "yyyy/mm/dd HH:MM")
                On Error GoTo ErrHandler
                
                lstFiles.AddItem fileName
                rowIdx = lstFiles.ListCount - 1
                lstFiles.List(rowIdx, 1) = fileSize
                lstFiles.List(rowIdx, 2) = modDate
                fileCount = fileCount + 1
            End If
        End If
        fileName = Dir()
    Loop
    
    ' 更新件数标签
    If fileCount = 0 Then
        lblCount.Caption = "選択件数：0 件（ファイルなし）"
    Else
        lblCount.Caption = "選択件数：0 件（合計 " & fileCount & " 件）"
    End If
    
    chkSelectAll.Value = False
    Exit Sub

ErrHandler:
    lstFiles.Clear
    lblCount.Caption = "選択件数：0 件（読込エラー）"
End Sub
```

---

### 修改 txtFolderPath_Change 事件，使用安全版函数

**找到：**

```vba
Private Sub txtFolderPath_Change()
    Call RefreshFileList
End Sub
```

**替换为：**

```vba
Private Sub txtFolderPath_Change()
    Call SafeRefreshFileList
End Sub
```

---

## 修改 2：添加欢迎界面（frmWelcome）

### 2-A：新建 UserForm

1. 在 VBA 编辑器中，点击菜单栏 **插入 → 用户窗体（UserForm）**
2. 在右侧属性窗口（按 F4 打开）设置以下属性：

| 属性 | 值 |
|------|-----|
| `(Name)` | `frmWelcome` |
| `Caption` | `勤怠データ集計ツール` |
| `Width` | `480` |
| `Height` | `320` |
| `StartUpPosition` | `2 - スクリーンの中央` |

### 2-B：在 frmWelcome 上添加控件

按以下表格添加控件（工具箱中选择对应类型后拖拽到窗体上）：

| 控件类型 | Name | Caption / Text | Left | Top | Width | Height | Font Size |
|---------|------|----------------|------|-----|-------|--------|-----------|
| Label | `lblTitle` | `勤怠データ集計ツール` | 40 | 30 | 400 | 40 | 18（加粗）|
| Label | `lblVersion` | `Version 3` | 40 | 75 | 400 | 20 | 10 |
| Label | `lblDesc` | `勤怠表から作業時間を集計し、請求書を自動生成します。` | 40 | 110 | 400 | 40 | 10 |
| Label | `lblDesc2` | `「開始」ボタンをクリックしてください。` | 40 | 155 | 400 | 20 | 10 |
| CommandButton | `btnStart` | `開始（Start）` | 100 | 210 | 120 | 36 | 11（加粗）|
| CommandButton | `btnExit` | `終了（Exit）` | 260 | 210 | 120 | 36 | 11 |

> **提示**：`lblTitle` 字体设置：右键控件 → 属性 → Font → 勾选 Bold，Size 填 18。

### 2-C：填写 frmWelcome 代码

在项目树中双击 **frmWelcome**，然后按 F7 打开代码窗口，**粘贴以下代码**（清空所有原有内容后粘贴）：

```vba
Option Explicit

'--------------------------------------------------
' 「開始」ボタン：メイン画面を表示
'--------------------------------------------------
Private Sub btnStart_Click()
    Me.Hide
    frmMain.Show
End Sub

'--------------------------------------------------
' 「終了」ボタン：ツールを閉じる
'--------------------------------------------------
Private Sub btnExit_Click()
    Dim answer As Integer
    answer = MsgBox("ツールを終了しますか？", vbYesNo + vbQuestion, "確認")
    If answer = vbYes Then
        Application.ScreenUpdating = True
        ThisWorkbook.Close SaveChanges:=False
    End If
End Sub

'--------------------------------------------------
' 右上 × ボタンで閉じるとき
'--------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Dim answer As Integer
        answer = MsgBox("ツールを終了しますか？", vbYesNo + vbQuestion, "確認")
        If answer = vbYes Then
            Application.ScreenUpdating = True
            ThisWorkbook.Close SaveChanges:=False
        Else
            Cancel = True
        End If
    End If
End Sub
```

### 2-D：修改 ThisWorkbook 模块

在项目树中双击 **ThisWorkbook**，找到 `Workbook_Open` 事件：

**找到这段原始代码：**

```vba
Private Sub Workbook_Open()
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Me.Worksheets(SHEET_SIMP_TRAD)
    If Not ws Is Nothing Then
        ws.Visible = xlSheetHidden
    End If
    On Error GoTo 0
    
    On Error Resume Next
    Set ws = Me.Worksheets(SHEET_SETTINGS)
    If Not ws Is Nothing Then
        ws.Visible = xlSheetHidden
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    ' メイン画面を表示
    frmMain.Show
End Sub
```

**替换为（只修改最后一行 `frmMain.Show` → `frmWelcome.Show`）：**

```vba
Private Sub Workbook_Open()
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Me.Worksheets(SHEET_SIMP_TRAD)
    If Not ws Is Nothing Then
        ws.Visible = xlSheetHidden
    End If
    On Error GoTo 0
    
    On Error Resume Next
    Set ws = Me.Worksheets(SHEET_SETTINGS)
    If Not ws Is Nothing Then
        ws.Visible = xlSheetHidden
    End If
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    
    ' まず歓迎画面を表示（v3改善）
    frmWelcome.Show
End Sub
```

---

## 修改 3：业务期间自动填写（不需要手动输入）

> **方案说明**：DTPicker 控件需要额外安装 OCX 组件，在某些 Windows 版本上可能不可用。  
> 本方案**无需安装任何控件**，改为在窗体初始化时自动填写当月第一天和最后一天，  
> 用户只需在确有必要时手动修改，大幅减少手动输入的需求。
>
> 如果需要完整的日历控件方案，请参考本文档末尾的「附录：DTPicker 安装说明」。

### 修改 UserForm_Initialize，自动设置当月日期

在 frmMain 代码中，找到 `UserForm_Initialize` 函数（约第 19 行），在函数末尾（`End Sub` 之前）添加以下代码：

**找到：**

```vba
    ' ファイル一覧を更新
    If Len(txtFolderPath.Text) > 0 Then
        Call RefreshFileList
    End If
End Sub
```

**替换为：**

```vba
    ' 業務期間を当月の初日と末日で自動設定（v3改善）
    Dim firstDay As String
    Dim lastDay As String
    firstDay = Format(DateSerial(Year(Date), Month(Date), 1), "yyyymmdd")
    lastDay = Format(DateSerial(Year(Date), Month(Date) + 1, 0), "yyyymmdd")
    
    ' 既に入力済みの場合は上書きしない
    If Len(Trim(txtPeriod.Text)) = 0 Then
        txtPeriod.Text = firstDay
    End If
    If Len(Trim(txtPeriodTo.Text)) = 0 Then
        txtPeriodTo.Text = lastDay
    End If
    
    ' ファイル一覧を更新
    If Len(txtFolderPath.Text) > 0 Then
        Call SafeRefreshFileList
    End If
End Sub
```

### （オプション）在期间文本框旁边添加「当月」快捷按钮

如果希望添加按钮快速设置日期，请在 **frmMain 的设计视图**中：

1. 在 `txtPeriod` 旁边添加一个 `CommandButton`，设置：
   - `Name`: `btnThisMonth`
   - `Caption`: `当月`
   - `Width`: 42, `Height`: 18

2. 在代码窗口添加以下事件：

```vba
'--------------------------------------------------
' 「当月」ボタン：業務期間を今月の初日・末日に設定
'--------------------------------------------------
Private Sub btnThisMonth_Click()
    txtPeriod.Text = Format(DateSerial(Year(Date), Month(Date), 1), "yyyymmdd")
    txtPeriodTo.Text = Format(DateSerial(Year(Date), Month(Date) + 1, 0), "yyyymmdd")
End Sub
```

---

## 修改完成后的测试步骤

1. **保存**：在 VBA 编辑器中按 **Ctrl+S**，选择"保留宏功能"
2. **关闭** VBA 编辑器
3. **重新打开** `toolV2_v3.xlsm`，验证：
   - [ ] 出现欢迎界面（`frmWelcome`），显示工具标题
   - [ ] 点击「開始」按钮后，正确显示主界面（`frmMain`）
   - [ ] 主界面中，`txtPeriod` 和 `txtPeriodTo` 已自动填写当月日期
   - [ ] 从下拉框选择「大連」→ 文件路径自动填写，文件列表正常显示（不报错）
   - [ ] 如果文件夹为空或不存在，列表框清空并显示友好提示
4. **验证集计功能**：选择文件，执行集计，确认原有功能正常

---

## 修改清单汇总

| # | 修改位置 | 修改内容 | 完成 |
|---|---------|---------|------|
| 1a | `frmMain` → `cboLocation_Change()` | 替换为含自动路径设置和错误处理的版本 | ☐ |
| 1b | `frmMain` → 新增 `SafeRefreshFileList()` | 添加完整容错逻辑的安全版文件扫描函数 | ☐ |
| 1c | `frmMain` → `txtFolderPath_Change()` | 改为调用 `SafeRefreshFileList` | ☐ |
| 2a | 新建 UserForm `frmWelcome` | 创建欢迎界面，添加5个控件 | ☐ |
| 2b | `frmWelcome` 代码 | 添加開始/終了按钮事件 | ☐ |
| 2c | `ThisWorkbook` → `Workbook_Open()` | 最后一行改为 `frmWelcome.Show` | ☐ |
| 3a | `frmMain` → `UserForm_Initialize()` | 添加自动填写当月日期的代码 | ☐ |
| 3b | `frmMain`（可选）| 添加「当月」快捷按钮 | ☐ |

---

## 附录：DTPicker 日历控件安装说明（可选方案）

> 如果需要完整的图形化日历选择器，可按以下步骤安装。  
> **注意**：此功能需要管理员权限，且在 64-bit Office 上可能不可用。

### 安装步骤

1. 在 VBA 编辑器菜单中点击 **工具 → 引用（References）**
2. 在列表中找到 **Microsoft Date and Time Picker Control 6.0 (SP6)**
3. 勾选后点击 OK

如果找不到该控件：

1. 打开命令提示符（管理员模式）
2. 执行：`regsvr32 "C:\Windows\SysWOW64\MSCOMCT2.OCX"`
3. 重新打开 VBA 编辑器并重试

### 使用 DTPicker 替换文本框（已安装控件后）

安装成功后，在 frmMain 设计视图中：

1. 工具箱 → 右键 → **附加控件（Additional Controls）**
2. 勾选 **Microsoft Date and Time Picker Control 6.0**
3. 删除 `txtPeriod` 和 `txtPeriodTo` 控件
4. 从工具箱拖放两个 DTPicker 控件，命名为 `dtpStart` 和 `dtpEnd`
5. 在 `UserForm_Initialize()` 中替换日期设置代码：

```vba
' DTPicker 版本
dtpStart.Value = DateSerial(Year(Date), Month(Date), 1)
dtpEnd.Value = DateSerial(Year(Date), Month(Date) + 1, 0)
dtpStart.Format = dtpCustom
dtpStart.CustomFormat = "yyyy/MM/dd"
dtpEnd.Format = dtpCustom
dtpEnd.CustomFormat = "yyyy/MM/dd"
```

6. 在 `btnExecute_Click()` 中，将原来读取 `txtPeriod.Text` 的地方改为：

```vba
Dim period As String
period = Format(dtpStart.Value, "yyyymmdd")
Dim periodTo As String  
periodTo = Format(dtpEnd.Value, "yyyymmdd")
```

---

*本手順書作成日：2026年5月*  
*対象ファイル：toolV2.xlsm → toolV2_v3.xlsm*
