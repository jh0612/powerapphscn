# Excel VBA 勤怠データ集計・請求書自動生成ツール 作成手順書（完全初心者向け・日本語環境）

## はじめに

### このツールでできること

このツールは、複数拠点（北京・大連・河南）から集まった勤怠表（Excelファイル）を自動で集計し、以下の3種類の請求書を一括生成するVBAマクロです。

1. **HSCN验收书**（検収書）：日別の勤怠明細と人員一覧を埋め込んだ検収書
2. **会社明細**（会社別明細書）：人員ごとの作業期間・工数・人月を一覧化した明細書
3. **实际报价兼发注书**（見積兼発注書）：単価・工数・金額を記載した発注書

### PowerApps版との違い

| 項目 | PowerApps版 | Excel VBA版（本ツール） |
|------|-------------|------------------------|
| 実行環境 | PowerApps（クラウド） | Excel（デスクトップ） |
| インターネット接続 | 必須 | 不要 |
| ライセンス費用 | Microsoft 365 ライセンスが必要 | Excelが使えれば無料 |
| データ保存場所 | SharePoint / Dataverse | ローカルフォルダ |
| カスタマイズ性 | 低（GUIベース） | 高（コード自由記述） |
| 処理速度 | クラウド依存 | ローカルPC性能依存（高速） |
| オフライン利用 | 不可 | 可 |

### 完成イメージ

起動するとメイン画面が表示されます。拠点フォルダを選択し、担当者を選び、集計したい勤怠ファイルにチェックを入れて「勤怠集計実行」ボタンを押すだけで、勤務時間が自動計算されます。その後「3ファイル一括生成」ボタンで請求書3点が自動出力されます。

---

## 目次

1. [前提準備](#1-前提準備)
2. [Step 1：フォルダ構成の準備](#step-1フォルダ構成の準備)
3. [Step 2：テンプレートファイルの準備](#step-2テンプレートファイルの準備hscn验收书会社明細实际报价兼发注書)
4. [Step 3：マクロ有効ブックの作成](#step-3マクロ有効ブックの作成)
5. [Step 4：設定シート・简繁对照シートの作成](#step-4設定シート简繁对照シートの作成)
6. [Step 5：UserForm frmMain（メイン画面）の作成](#step-5userform-frmmainメイン画面の作成)
7. [Step 6：UserForm frmProject（プロジェクト設定画面）の作成](#step-6userform-frmprojectプロジェクト設定画面の作成)
8. [Step 7：UserForm frmPerson（人員設定画面）の作成](#step-7userform-frmperson人員設定画面の作成)
9. [Step 8：UserForm frmPrice（単価表設定画面）の作成](#step-8userform-frmprice単価表設定画面の作成)
10. [Step 9：標準モジュール modConstants のコード](#step-9標準モジュール-modconstants-のコード)
11. [Step 10：標準モジュール modUtils のコード](#step-10標準モジュール-modutils-のコード)
12. [Step 11：標準モジュール modMain のコード](#step-11標準モジュール-modmain-のコード)
13. [Step 12：frmMain のコード](#step-12frmmain-のコード)
14. [Step 13：frmProject のコード](#step-13frmproject-のコード)
15. [Step 14：frmPerson のコード](#step-14frmperson-のコード)
16. [Step 15：frmPrice のコード](#step-15frmprice-のコード)
17. [Step 16：ThisWorkbook のコード](#step-16thisworkbook-のコード)
18. [Step 17：動作テスト](#step-17動作テスト)
19. [補足：トラブルシューティング FAQ](#補足トラブルシューティング-faq)

---

## 1. 前提準備

**このステップでやること**：ツール作成に必要な環境とファイルを確認します。

### 必要なもの

| 項目 | 説明 |
|------|------|
| Microsoft Excel | 2016以降（Windows版）。Mac版は非対応です |
| 勤怠表ファイル | 各拠点（北京/大連/河南）の勤怠表Excelファイル。拠点別フォルダに格納されている前提 |
| 単価表ファイル | 氏名・等級・契約単価が記載されたExcelファイル |
| テンプレートファイル | HSCN验收书・会社明細・实际报价兼发注书の3点（既存のExcelファイル） |
| 簡繁体変換表 | 簡体字→繁体字のマッピング（本手順書にプリセットとして収録） |

> **重要**：本手順書はWindows 11 + Excel 2016/2019/365 日本語版での操作を前提としています。

---

## Step 1：フォルダ構成の準備

**このステップでやること**：ツールとデータを格納するフォルダ構成を作ります。

### 推奨フォルダ構成

```
C:\Work\勤怠集計\
├── Tool\                    ← マクロ有効ブック（本ツール）を置く
├── Data\                    ← 勤怠表ファイルを置く
│   ├── 北京\
│   ├── 大連\
│   └── 河南\
├── Template\                ← 3種類のテンプレートファイルを置く
│   ├── HSCN验收书_Template.xlsx
│   ├── 会社明細_Template.xlsx
│   └── 实际报价兼发注书_Template.xlsx
├── 単価表\                  ← 単価表ファイルを置く
│   └── 単価表_2026.xlsx
└── Output\                  ← 出力先（ツールが自動生成）
```

### 操作手順

1. エクスプローラーを開き、`C:\Work\勤怠集計` フォルダを作成します
2. その中に上記のサブフォルダ（Tool、Data、Template、単価表、Output）をすべて作成します
3. Dataフォルダの中に「北京」「大連」「河南」フォルダを作成します
4. 既存の勤怠表ファイルを各拠点フォルダに配置します
5. テンプレートファイル3点をTemplateフォルダに配置します
6. 単価表ファイルを単価表フォルダに配置します

> **重要**：フォルダ構成は任意ですが、上記の構成にすることで後の設定がスムーズです。別の構成でもツールは動作します（設定画面でパスを指定するため）。

---

## Step 2：テンプレートファイルの準備（HSCN验收书/会社明細/实际报价兼发注书）

**このステップでやること**：出力の雛形となる3種類のテンプレートExcelファイルを準備します。

### 2-1. HSCN验收书 テンプレート

以下のシート構成を持つExcelファイルを用意します。

#### シート構成

| シート名 | 内容 |
|----------|------|
| Acceptance Request | 検収依頼書（表紙）。Row 5〜12：プロジェクト情報、Row 16以降：人員別集計データ |
| Timesheet_Template | 日別勤怠明細の雛形シート（ツールがコピーして `Timesheet_{氏名}` を作成） |

#### Acceptance Requestシートのセル配置

```
Row 1: タイトル「HSCN验收书 / Acceptance Request」
Row 3: （空行）
Row 5: 项目名称（プロジェクト名）
Row 6: 供应商名称（サプライヤー名）
Row 7: 业务形式（業務形式）
Row 8: 作业内容（作業内容）
Row 9: 成果物
Row 10: 阶段目标（マイルストーン）
Row 11: 合同编号（契約番号）
Row 12: 采购订单号（発注番号）
Row 14: （空行）
Row 15: ヘッダー行「No / 氏名 / 作業期間 / 作業内容 / 総工時(h) / 人月 / 等級 / 契約単価 / 金額」
Row 16〜: データ行（ツールが自動埋め込み）
```

#### Timesheet_Templateシートのセル配置

```
Row 1: タイトル「HSCN验收书 - Timesheet」
Row 3: （空行）
Row 4: 会社名（Company Name）              ← ツールが自動埋め込み
Row 5: （空行）
Row 6: 氏名(Name) / 等級(Level)           ← ツールが自動埋め込み
Row 7: （空行）
Row 8: 年月(YYYY/MM) / 総工時(Total Hrs) / 人月(Man-Month)  ← ツールが自動埋め込み
Row 9: （空行）
Row 10: ヘッダー行「月 / 日 / 作業内容 / 上班时间 / 下班时间 / 加班(h) / 缺勤(h) / 工时(h)」
Row 11〜: データ行（ツールが自動埋め込み）
Row 11以降の最終行の下: 合計行（テンプレート側の計算式で自動計算）
```

### 2-2. 会社明細 テンプレート

| 項目 | 内容 |
|------|------|
| シート名 | Sheet1（任意） |
| Row 1 | ヘッダー行「No / 氏名 / 作業期間 / 作業内容 / 総工時(h) / 人月 / 等級 / 契約単価 / 金額」 |
| Row 2〜 | データ行（ツールが自動埋め込み） |

Row 2以降はツールが自動でデータを書き込みます。合計行はテンプレート側で計算式（SUM関数など）を事前に設定しておきます。

### 2-3. 实际报价兼发注书 テンプレート

| 項目 | 内容 |
|------|------|
| シート名 | Sheet1（任意） |
| Row 1〜12 | ヘッダー部（プロジェクト情報、発注者情報など。手動で入力済みの前提） |
| Row 13 | 表一のヘッダー行「No / 氏名 / 作業期間 / 作業内容 / 総工時(h) / 人月 / 等級 / 契約単価(円) / 金額(円)」 |
| Row 14〜 | 表一のデータ行（ツールが自動埋め込み） |

> **重要**：テンプレート内の計算式（合計金額、人月の自動計算など）はテンプレート側で事前に設定しておいてください。ツールは生データ（氏名、工数、単価など）のみを埋め込みます。

---

## Step 3：マクロ有効ブックの作成

**このステップでやること**：VBAマクロを格納するExcelブック（.xlsm）を作成します。

### 操作手順

1. Excelを起動します
2. 「空白のブック」をクリックして新規ブックを作成します
3. キーボードの `Alt` キーを押して離し、続いて `F` キー、`A` キーの順に押します（または「ファイル」→「名前を付けて保存」）
4. 保存先に `C:\Work\勤怠集計\Tool\` を選択します
5. ファイル名に `勤怠集計ツール` と入力します
6. 「ファイルの種類」で **「Excel マクロ有効ブック (*.xlsm)」** を選択します
7. 「保存」をクリックします

### リボンに開発タブを表示する

VBAを使うには「開発」タブが必要です。表示されていない場合は以下の手順で追加します。

1. 「ファイル」タブ → 「オプション」をクリック
2. 左メニューから「リボンのユーザー設定」を選択
3. 右側の「メイン タブ」一覧で「開発」にチェックを入れる
4. 「OK」をクリック

### VBE（Visual Basic Editor）の起動方法

1. 「開発」タブをクリック
2. 「Visual Basic」ボタンをクリック（または `Alt` + `F11` キー）

VBE画面の構成：

```
┌─────────────────────────────────────────────────────┐
│ メニューバー                                        │
├──────────┬──────────────────────────────────────────┤
│プロジェクト│                                          │
│エクスプロ  │          コード編集エリア                  │
│ーラー     │        （メイン作業領域）                   │
│           │                                          │
│ VBAProject│                                          │
│ ├─Sheet1  │                                          │
│ ├─Sheet2  │                                          │
│ ├─Sheet3  │                                          │
│ └─ThisWork│                                          │
│           │                                          │
├──────────┴──────────────────────────────────────────┤
│ プロパティウィンドウ / イミディエイトウィンドウ         │
└─────────────────────────────────────────────────────┘
```

- **左上**：プロジェクトエクスプローラー（ブック内の全オブジェクト一覧）
- **左下**：プロパティウィンドウ（選択したオブジェクトのプロパティ設定）
- **右側**：コード編集エリア（VBAコードを記述する場所）

---

## Step 4：設定シート・简繁对照シートの作成

**このステップでやること**：設定情報を保存するシートと、簡体字→繁体字の変換表を作成します。

### 4-1. 設定シート（Settings）の作成

1. Excelに戻り、Sheet1を右クリック → 「名前の変更」で **「Settings」** に変更します
2. 以下の表のようにデータを入力します（列A：設定項目名、列B：設定値）

| A（設定項目） | B（設定値） |
|---------------|-------------|
| ProjectName | （空欄） |
| Supplier | （空欄） |
| ServiceMode | T&M |
| ServiceContent | （空欄） |
| Deliverable | （空欄） |
| Milestone | （空欄） |
| ContractNo | （空欄） |
| PO | （空欄） |
| OrderNo | （空欄） |
| HSCNPath | （空欄） |
| CompanyPath | （空欄） |
| OrderPath | （空欄） |
| OutputPath | （空欄） |
| PriceFilePath | （空欄） |
| PriceSheetName | Sheet1 |
| PriceHeaderRow | 1 |
| PriceNameCol | A |
| PriceEndCol | B |
| PriceUnitCol | C |

3. 列Aの幅を広げて見やすくします（列Aの右端をダブルクリック）

### 4-2. 简繁对照シートの作成

1. Sheet2を右クリック → 「名前の変更」で **「简繁对照」** に変更します
2. 列Aに簡体字、列Bに対応する繁体字を入力します
3. 以下のプリセットデータをコピーして貼り付けます（代表的な100組）

| A（簡体字） | B（繁体字） |
|-------------|-------------|
| 张 | 張 |
| 刘 | 劉 |
| 陈 | 陳 |
| 杨 | 楊 |
| 赵 | 趙 |
| 黄 | 黃 |
| 周 | 周 |
| 吴 | 吳 |
| 徐 | 徐 |
| 孙 | 孫 |
| 马 | 馬 |
| 朱 | 朱 |
| 胡 | 胡 |
| 郭 | 郭 |
| 何 | 何 |
| 高 | 高 |
| 林 | 林 |
| 罗 | 羅 |
| 梁 | 梁 |
| 谢 | 謝 |
| 宋 | 宋 |
| 唐 | 唐 |
| 许 | 許 |
| 韩 | 韓 |
| 冯 | 馮 |
| 邓 | 鄧 |
| 曹 | 曹 |
| 彭 | 彭 |
| 曾 | 曾 |
| 肖 | 肖 |
| 田 | 田 |
| 董 | 董 |
| 袁 | 袁 |
| 潘 | 潘 |
| 于 | 于 |
| 蒋 | 蔣 |
| 蔡 | 蔡 |
| 余 | 余 |
| 杜 | 杜 |
| 叶 | 葉 |
| 程 | 程 |
| 苏 | 蘇 |
| 魏 | 魏 |
| 吕 | 呂 |
| 丁 | 丁 |
| 任 | 任 |
| 沈 | 沈 |
| 姚 | 姚 |
| 卢 | 盧 |
| 姜 | 姜 |
| 崔 | 崔 |
| 钟 | 鍾 |
| 谭 | 譚 |
| 陆 | 陸 |
| 汪 | 汪 |
| 范 | 范 |
| 金 | 金 |
| 石 | 石 |
| 廖 | 廖 |
| 贾 | 賈 |
| 夏 | 夏 |
| 韦 | 韋 |
| 付 | 付 |
| 方 | 方 |
| 白 | 白 |
| 邹 | 鄒 |
| 孟 | 孟 |
| 熊 | 熊 |
| 秦 | 秦 |
| 邱 | 邱 |
| 江 | 江 |
| 尹 | 尹 |
| 薛 | 薛 |
| 闫 | 閆 |
| 段 | 段 |
| 雷 | 雷 |
| 侯 | 侯 |
| 龙 | 龍 |
| 史 | 史 |
| 陶 | 陶 |
| 黎 | 黎 |
| 贺 | 賀 |
| 顾 | 顧 |
| 毛 | 毛 |
| 郝 | 郝 |
| 龚 | 龔 |
| 邵 | 邵 |
| 万 | 萬 |
| 钱 | 錢 |
| 严 | 嚴 |
| 覃 | 覃 |
| 武 | 武 |
| 戴 | 戴 |
| 莫 | 莫 |
| 孔 | 孔 |
| 向 | 向 |
| 汤 | 湯 |
| 温 | 溫 |
| 康 | 康 |
| 施 | 施 |
| 文 | 文 |
| 牛 | 牛 |
| 樊 | 樊 |

4. 1行目に「簡体字」「繁体字」という見出しを入れておきます（任意）

> **重要**：简繁对照シートはツールが内部で使用します。非表示にしても問題ありません（ツールが自動で読み取ります）。シートの非表示は、シート名を右クリック → 「表示しない」で設定できます。

### 4-3. Sheet3の削除

使用しないSheet3は削除します。Sheet3を右クリック → 「削除」を選択します。

### 保存

ここまでの作業を保存します（`Ctrl` + `S`）。

---

## Step 5：UserForm frmMain（メイン画面）の作成

**このステップでやること**：ツールのメイン操作画面（UserForm）を作成します。

### 5-1. UserFormの挿入

1. VBEを開きます（`Alt` + `F11`）
2. メニューから「挿入」→「ユーザーフォーム」をクリック
3. プロジェクトエクスプローラーに「UserForm1」が追加されます

### 5-2. UserFormのプロパティ設定

プロパティウィンドウ（左下）で以下の値を設定します。プロパティウィンドウが表示されていない場合は、メニュー「表示」→「プロパティウィンドウ」で表示します。

| プロパティ | 設定値 |
|------------|--------|
| **(オブジェクト名)** | **frmMain** |
| Caption | 勤怠データ集計・請求書自動生成ツール |
| Width | 720 |
| Height | 620 |
| StartUpPosition | 2 - 画面の中央 |

> **重要**：**(オブジェクト名)** はプロパティウィンドウの一番上にあります。名前の先頭に `(オブジェクト名)` と表示されています。ここを `frmMain` に変更してください。`Caption`（キャプション）とは別のプロパティです。

### 5-3. コントロールの配置

ツールボックスが表示されていない場合は、メニュー「表示」→「ツールボックス」で表示します。

#### コントロール配置図

```
┌──────────────────────────────────────────────────────────────────┐
│ 勤怠データ集計・請求書自動生成ツール                               │
├──────────────────────────────────────────────────────────────────┤
│ 拠点選択：[cboLocation ▼]  フォルダパス：[txtFolderPath        ] [参照...] │
│ 担当者名：[cboPerson ▼]    業務期間：[txtPeriod                ]          │
│ ☑ 全选                                             選択件数：0 件       │
│ ┌──────────────────────────────────────────────────────────────┐ │
│ │ ファイル名                     │ サイズ    │ 更新日時        │ │
│ │ 北京_202604_勤怠.xlsx          │ 45 KB     │ 2026/04/30      │ │
│ │ 北京_202605_勤怠.xlsx          │ 52 KB     │ 2026/05/31      │ │
│ │ ...                            │ ...       │ ...             │ │
│ └──────────────────────────────────────────────────────────────┘ │
│ [プロジェクト設定] [人員設定] [単価表設定]                         │
│ [勤怠集計実行] [3ファイル一括生成] [出力フォルダを開く]             │
│ ステータス：待機中                                                 │
│ 処理履歴：                                                         │
│ ┌──────────────────────────────────────────────────────────────┐ │
│ │ 2026/05/27 10:00 - 集計完了：3件                              │ │
│ └──────────────────────────────────────────────────────────────┘ │
└──────────────────────────────────────────────────────────────────┘
```

#### コントロール一覧とプロパティ

ツールボックスから以下のコントロールをドラッグ＆ドロップで配置し、プロパティを設定します。

**1行目：拠点選択**

| コントロール | 種類 | プロパティ | 設定値 |
|-------------|------|------------|--------|
| ラベル1 | Label | **(オブジェクト名)** | lblLocation |
| | | Caption | 拠点選択： |
| | | Top | 12 |
| | | Left | 12 |
| | | Width | 60 |
| コンボボックス1 | ComboBox | **(オブジェクト名)** | cboLocation |
| | | Top | 10 |
| | | Left | 72 |
| | | Width | 100 |
| ラベル2 | Label | **(オブジェクト名)** | lblFolderPath |
| | | Caption | フォルダパス： |
| | | Top | 12 |
| | | Left | 190 |
| | | Width | 72 |
| テキストボックス1 | TextBox | **(オブジェクト名)** | txtFolderPath |
| | | Top | 10 |
| | | Left | 262 |
| | | Width | 280 |
| コマンドボタン1 | CommandButton | **(オブジェクト名)** | btnBrowse |
| | | Caption | 参照... |
| | | Top | 8 |
| | | Left | 550 |
| | | Width | 60 |
| | | Height | 24 |

**2行目：担当者・業務期間**

| コントロール | 種類 | プロパティ | 設定値 |
|-------------|------|------------|--------|
| ラベル3 | Label | **(オブジェクト名)** | lblPerson |
| | | Caption | 担当者名： |
| | | Top | 44 |
| | | Left | 12 |
| | | Width | 60 |
| コンボボックス2 | ComboBox | **(オブジェクト名)** | cboPerson |
| | | Top | 42 |
| | | Left | 72 |
| | | Width | 150 |
| ラベル4 | Label | **(オブジェクト名)** | lblPeriod |
| | | Caption | 業務期間： |
| | | Top | 44 |
| | | Left | 240 |
| | | Width | 60 |
| テキストボックス2 | TextBox | **(オブジェクト名)** | txtPeriod |
| | | Top | 42 |
| | | Left | 300 |
| | | Width | 200 |

**3行目：全选チェック・件数**

| コントロール | 種類 | プロパティ | 設定値 |
|-------------|------|------------|--------|
| チェックボックス1 | CheckBox | **(オブジェクト名)** | chkSelectAll |
| | | Caption | 全选 |
| | | Top | 76 |
| | | Left | 12 |
| | | Width | 60 |
| ラベル5 | Label | **(オブジェクト名)** | lblCount |
| | | Caption | 選択件数：0 件 |
| | | Top | 76 |
| | | Left | 80 |
| | | Width | 200 |

**4行目：ファイルリスト**

| コントロール | 種類 | プロパティ | 設定値 |
|-------------|------|------------|--------|
| リストボックス1 | ListBox | **(オブジェクト名)** | lstFiles |
| | | Top | 100 |
| | | Left | 12 |
| | | Width | 690 |
| | | Height | 200 |
| | | ListStyle | 1 - fmListStyleOption |
| | | MultiSelect | 1 - fmMultiSelectMulti |
| | | ColumnCount | 3 |
| | | ColumnWidths | 300;120;200 |

**5行目：設定ボタン**

| コントロール | 種類 | プロパティ | 設定値 |
|-------------|------|------------|--------|
| コマンドボタン2 | CommandButton | **(オブジェクト名)** | btnProject |
| | | Caption | プロジェクト設定 |
| | | Top | 312 |
| | | Left | 12 |
| | | Width | 120 |
| | | Height | 28 |
| コマンドボタン3 | CommandButton | **(オブジェクト名)** | btnPerson |
| | | Caption | 人員設定 |
| | | Top | 312 |
| | | Left | 144 |
| | | Width | 100 |
| | | Height | 28 |
| コマンドボタン4 | CommandButton | **(オブジェクト名)** | btnPrice |
| | | Caption | 単価表設定 |
| | | Top | 312 |
| | | Left | 256 |
| | | Width | 100 |
| | | Height | 28 |

**6行目：実行ボタン**

| コントロール | 種類 | プロパティ | 設定値 |
|-------------|------|------------|--------|
| コマンドボタン5 | CommandButton | **(オブジェクト名)** | btnExecute |
| | | Caption | 勤怠集計実行 |
| | | Top | 350 |
| | | Left | 12 |
| | | Width | 140 |
| | | Height | 32 |
| | | BackColor | &H00C0FFC0&（薄緑） |
| コマンドボタン6 | CommandButton | **(オブジェクト名)** | btnGenerate |
| | | Caption | 3ファイル一括生成 |
| | | Top | 350 |
| | | Left | 164 |
| | | Width | 150 |
| | | Height | 32 |
| | | BackColor | &H00C0C0FF&（薄赤） |
| コマンドボタン7 | CommandButton | **(オブジェクト名)** | btnOpenFolder |
| | | Caption | 出力フォルダを開く |
| | | Top | 350 |
| | | Left | 326 |
| | | Width | 150 |
| | | Height | 32 |

**7行目：ステータス**

| コントロール | 種類 | プロパティ | 設定値 |
|-------------|------|------------|--------|
| ラベル6 | Label | **(オブジェクト名)** | lblStatus |
| | | Caption | 待機中 |
| | | Top | 394 |
| | | Left | 12 |
| | | Width | 690 |
| | | Height | 20 |
| | | ForeColor | &H00800000&（濃青） |

**8行目：処理履歴**

| コントロール | 種類 | プロパティ | 設定値 |
|-------------|------|------------|--------|
| ラベル7 | Label | **(オブジェクト名)** | lblHistory |
| | | Caption | 処理履歴： |
| | | Top | 418 |
| | | Left | 12 |
| | | Width | 80 |
| リストボックス2 | ListBox | **(オブジェクト名)** | lstHistory |
| | | Top | 438 |
| | | Left | 12 |
| | | Width | 690 |
| | | Height | 60 |

### 5-4. TabIndexの調整

`Tab` キーでの移動順序を調整します。各コントロールを配置した順に `TabIndex` が割り当てられますが、必要に応じて以下の順に調整してください。

1. cboLocation (0)
2. txtFolderPath (1)
3. btnBrowse (2)
4. cboPerson (3)
5. txtPeriod (4)
6. chkSelectAll (5)
7. lstFiles (6)
8. btnProject (7)
9. btnPerson (8)
10. btnPrice (9)
11. btnExecute (10)
12. btnGenerate (11)
13. btnOpenFolder (12)
14. lstHistory (13)

TabIndexの変更方法：コントロールを選択 → プロパティウィンドウで `TabIndex` の値を変更します。

---

## Step 6：UserForm frmProject（プロジェクト設定画面）の作成

**このステップでやること**：プロジェクト情報とテンプレートパスを設定する画面を作成します。

### 6-1. UserFormの挿入と設定

1. VBEで「挿入」→「ユーザーフォーム」をクリック
2. プロパティを設定

| プロパティ | 設定値 |
|------------|--------|
| **(オブジェクト名)** | **frmProject** |
| Caption | プロジェクト基本情報・テンプレート設定 |
| Width | 700 |
| Height | 520 |
| StartUpPosition | 2 - 画面の中央 |

### 6-2. コントロールの配置

#### 配置図

```
┌──────────────────────────────────────────────────────────────────┐
│ プロジェクト基本情報・テンプレート設定                              │
├──────────────────────────────────────────────────────────────────┤
│ ── プロジェクト情報 ──                                           │
│ 項目名称：    [txtProjectName                              ]     │
│ 供应商名称：  [txtSupplier                                 ]     │
│ 业务形式：    [cboServiceMode ▼]                                  │
│ 作业内容：    [txtServiceContent                            ]     │
│ 成果物：      [txtDeliverable                               ]     │
│ 阶段目标：    [txtMilestone                                 ]     │
│ 合同编号：    [txtContractNo        ] 采购订单号：[txtPO         ] │
│ 発注番号：    [txtOrderNo           ]                             │
│                                                                   │
│ ── 出力テンプレートパス ──                                        │
│ HSCN验收书：  [txtHSCNPath          ] [参照...]                   │
│ 会社明細：    [txtCompanyPath       ] [参照...]                   │
│ 实际报价兼发注书：[txtOrderPath     ] [参照...]                   │
│ 出力先：      [txtOutputPath        ] [参照...]                   │
│                                                                   │
│              [保存して閉じる]   [キャンセル]                        │
└──────────────────────────────────────────────────────────────────┘
```

#### コントロール一覧とプロパティ

**セクションラベル1**

| コントロール | 種類 | プロパティ | 設定値 |
|-------------|------|------------|--------|
| ラベル1 | Label | Caption | ── プロジェクト情報 ── |
| | | Top | 10 |
| | | Left | 12 |
| | | Width | 300 |
| | | Font | Bold |

**プロジェクト情報（左ラベル + 右テキスト/コンボボックス）**

各ラベル：Left=12, Width=100, TextAlign=3 - fmTextAlignRight
各テキストボックス/コンボボックス：Left=120

| 行 | ラベル Caption | (オブジェクト名) | 種類 | Width | Top |
|----|---------|-----------------|------|-------|-----|
| 1 | 項目名称： | txtProjectName | TextBox | 400 | 40 |
| 2 | 供应商名称： | txtSupplier | TextBox | 400 | 72 |
| 3 | 业务形式： | cboServiceMode | ComboBox | 200 | 104 |
| 4 | 作业内容： | txtServiceContent | TextBox | 400 | 136 |
| 5 | 成果物： | txtDeliverable | TextBox | 400 | 168 |
| 6 | 阶段目标： | txtMilestone | TextBox | 400 | 200 |
| 7 | 合同编号： | txtContractNo | TextBox | 200 | 232 |
| 8 | 采购订单号： | txtPO | TextBox | 200 | 264 |
| 9 | 発注番号： | txtOrderNo | TextBox | 200 | 296 |

cboServiceMode：Style=2 - fmStyleDropDownList

**セクションラベル2**

| コントロール | 種類 | プロパティ | 設定値 |
|-------------|------|------------|--------|
| ラベル2 | Label | Caption | ── 出力テンプレートパス ── |
| | | Top | 334 |
| | | Left | 12 |
| | | Width | 300 |
| | | Font | Bold |

**テンプレートパス（ラベル + テキスト + 参照ボタン）**

各ラベル：Left=12, Width=130, TextAlign=3
各テキストボックス：Left=146
各参照ボタン：Caption="参照...", Width=60, Height=22, Left=554

| 行 | ラベル Caption | TextBox (オブジェクト名) | Width | ボタン (オブジェクト名) | Top |
|----|---------------|--------------------------|-------|------------------------|-----|
| 1 | HSCN验收书： | txtHSCNPath | 400 | btnHSCNBrowse | 364 |
| 2 | 会社明細： | txtCompanyPath | 400 | btnCompanyBrowse | 396 |
| 3 | 实际报价兼发注书： | txtOrderPath | 400 | btnOrderBrowse | 428 |
| 4 | 出力先： | txtOutputPath | 400 | btnOutputBrowse | 460 |

**下部ボタン**

| (オブジェクト名) | Caption | Top | Left | Width | Height |
|-----------------|---------|-----|------|-------|--------|
| btnSaveProject | 保存して閉じる | 476 | 460 | 110 | 28 |
| btnCancelProject | キャンセル | 476 | 580 | 90 | 28 |

---

## Step 7：UserForm frmPerson（人員設定画面）の作成

**このステップでやること**：作業人員の割り当てを登録する画面を作成します。

### 7-1. UserFormの挿入と設定

1. VBEで「挿入」→「ユーザーフォーム」をクリック
2. プロパティを設定

| プロパティ | 設定値 |
|------------|--------|
| **(オブジェクト名)** | **frmPerson** |
| Caption | 作業人員設定 |
| Width | 650 |
| Height | 500 |
| StartUpPosition | 2 - 画面の中央 |

### 7-2. コントロールの配置

#### 配置図

```
┌──────────────────────────────────────────────────────────────────┐
│ 作業人員設定                                                      │
├──────────────────────────────────────────────────────────────────┤
│ 氏名：[cboName ▼]  作業期間：[txtWorkPeriod    ]                 │
│ 作業内容：[txtWorkContent                             ]          │
│ 作業番号：[txtWorkNo      ]  等級：[cboLevel ▼]                  │
│                                                                   │
│ [追加] [更新] [削除] [全削除]                                     │
│                                                                   │
│ ┌──────────────────────────────────────────────────────────────┐ │
│ │ No │ 氏名   │ 作業期間        │ 作業内容   │ 作業番号 │ 等級   │ │
│ │ 1  │ 張三   │ 2026/4/1~4/20   │ 開発      │ W001     │ Senior1│ │
│ │ 2  │ 李四   │ 2026/4/1~4/30   │ テスト    │ W002     │ Junior │ │
│ └──────────────────────────────────────────────────────────────┘ │
│                                                                   │
│              [保存して閉じる]   [キャンセル]                        │
└──────────────────────────────────────────────────────────────────┘
```

#### コントロール一覧

**入力エリア**

| ラベル Caption | コントロール種類 | (オブジェクト名) | Width | Top | Left(ラベル) | Left(入力) |
|---------------|-----------------|-----------------|-------|-----|-------------|-----------|
| 氏名： | ComboBox | cboName | 150 | 12 | 12 | 72 |
| 作業期間： | TextBox | txtWorkPeriod | 200 | 12 | 240 | 310 |
| 作業内容： | TextBox | txtWorkContent | 400 | 44 | 12 | 72 |
| 作業番号： | TextBox | txtWorkNo | 120 | 76 | 12 | 72 |
| 等級： | ComboBox | cboLevel | 120 | 76 | 300 | 340 |

cboLevel にはプロパティ `Style` = `2 - fmStyleDropDownList` を設定。

**操作ボタン**

| (オブジェクト名) | Caption | Top | Left | Width | Height |
|-----------------|---------|-----|------|-------|--------|
| btnAdd | 追加 | 108 | 72 | 70 | 24 |
| btnUpdate | 更新 | 108 | 150 | 70 | 24 |
| btnDelete | 削除 | 108 | 228 | 70 | 24 |
| btnClearAll | 全削除 | 108 | 306 | 70 | 24 |

**リストボックス**

| プロパティ | 設定値 |
|------------|--------|
| **(オブジェクト名)** | lstPersonList |
| Top | 144 |
| Left | 12 |
| Width | 620 |
| Height | 220 |
| ColumnCount | 6 |
| ColumnWidths | 40;80;140;100;80;80 |
| ListStyle | 1 - fmListStyleOption |
| MultiSelect | 0 - fmMultiSelectSingle |

**下部ボタン**

| (オブジェクト名) | Caption | Top | Left | Width | Height |
|-----------------|---------|-----|------|-------|--------|
| btnSavePerson | 保存して閉じる | 376 | 410 | 110 | 28 |
| btnCancelPerson | キャンセル | 376 | 530 | 90 | 28 |

---

## Step 8：UserForm frmPrice（単価表設定画面）の作成

**このステップでやること**：単価表ファイルのパスと列位置を設定する画面を作成します。

### 8-1. UserFormの挿入と設定

1. VBEで「挿入」→「ユーザーフォーム」をクリック
2. プロパティを設定

| プロパティ | 設定値 |
|------------|--------|
| **(オブジェクト名)** | **frmPrice** |
| Caption | 単価表設定 |
| Width | 500 |
| Height | 350 |
| StartUpPosition | 2 - 画面の中央 |

### 8-2. コントロールの配置

#### 配置図

```
┌──────────────────────────────────────────────────────────────────┐
│ 単価表設定                                                        │
├──────────────────────────────────────────────────────────────────┤
│ 単価表ファイル：[txtPricePath                ] [参照...]          │
│ Sheet名：      [cboPriceSheet ▼]                                 │
│ 見出し行番号： [txtPriceHeaderRow]                                │
│ 氏名列：      [cboPriceNameCol ▼]    (A〜Z)                      │
│ 退社日列：    [cboPriceEndCol ▼]     (A〜Z)                      │
│ 契約単価列：  [cboPriceUnitCol ▼]    (A〜Z)                      │
│                                                                   │
│              [保存して閉じる]   [キャンセル]                        │
└──────────────────────────────────────────────────────────────────┘
```

#### コントロール一覧

| ラベル Caption | コントロール | (オブジェクト名) | Width | Top | Left(ラベル) | Left(入力) |
|---------------|-------------|-----------------|-------|-----|-------------|-----------|
| 単価表ファイル： | TextBox | txtPricePath | 300 | 14 | 12 | 110 |
| (参照ボタン) | CommandButton | btnPriceBrowse | 60 | 12 | — | 418 |
| Sheet名： | ComboBox | cboPriceSheet | 200 | 50 | 12 | 110 |
| 見出し行番号： | TextBox | txtPriceHeaderRow | 60 | 86 | 12 | 110 |
| 氏名列： | ComboBox | cboPriceNameCol | 80 | 122 | 12 | 110 |
| 退社日列： | ComboBox | cboPriceEndCol | 80 | 158 | 12 | 110 |
| 契約単価列： | ComboBox | cboPriceUnitCol | 80 | 194 | 12 | 110 |

btnPriceBrowse の Caption：「参照...」、Height: 22

**列選択コンボボックス（cboPriceNameCol, cboPriceEndCol, cboPriceUnitCol）**

共通プロパティ：Style = 2 - fmStyleDropDownList
選択肢（A〜Z）はコードで追加します。

**下部ボタン**

| (オブジェクト名) | Caption | Top | Left | Width | Height |
|-----------------|---------|-----|------|-------|--------|
| btnSavePrice | 保存して閉じる | 236 | 240 | 110 | 28 |
| btnCancelPrice | キャンセル | 236 | 360 | 90 | 28 |

---
## Step 9：標準モジュール modConstants のコード

**このステップでやること**：定数定義をまとめた標準モジュールを作成します。

### 9-1. 標準モジュールの挿入

1. VBEでメニュー「挿入」→「標準モジュール」をクリック
2. プロジェクトエクスプローラーに「Module1」が追加されます
3. プロパティウィンドウで **(オブジェクト名)** を **modConstants** に変更します

### 9-2. modConstants のコード

以下のコードを**すべてコピー**して、modConstantsのコード編集エリアに貼り付けてください。

```vba
Option Explicit

'==================================================
' 定数定義モジュール
' 勤怠集計ツールで使用する定数を一元管理します
'==================================================

' --- シート名定数 ---
Public Const SHEET_SETTINGS As String = "Settings"
Public Const SHEET_SIMP_TRAD As String = "简繁对照"

' --- 拠点定数 ---
Public Const LOCATION_BEIJING As String = "北京"
Public Const LOCATION_DALIAN As String = "大連"
Public Const LOCATION_HENAN As String = "河南"

' --- 勤務時間計算定数 ---
' 昼休憩時間（時間単位）
Public Const LUNCH_BREAK_HOURS As Double = 1#
' 標準勤務時間（時間単位）
Public Const STANDARD_WORK_HOURS As Double = 8#
' 1人月あたりの標準時間（168時間 = 21日 × 8時間）
Public Const HOURS_PER_MAN_MONTH As Double = 168#
' 勤務時間の丸め単位（0.5時間単位）
Public Const ROUND_UNIT_HOURS As Double = 0.5

' --- ファイル関連定数 ---
' テンプレートシート名
Public Const TEMPLATE_TIMESHEET As String = "Timesheet_Template"
Public Const TEMPLATE_ACCEPTANCE As String = "Acceptance Request"

' --- 出力ファイル名の接尾辞 ---
Public Const SUFFIX_TIMESHEET As String = "_Timesheet"

' --- 設定項目名（Settingsシートの列Aと一致） ---
Public Const SET_PROJECT_NAME As String = "ProjectName"
Public Const SET_SUPPLIER As String = "Supplier"
Public Const SET_SERVICE_MODE As String = "ServiceMode"
Public Const SET_SERVICE_CONTENT As String = "ServiceContent"
Public Const SET_DELIVERABLE As String = "Deliverable"
Public Const SET_MILESTONE As String = "Milestone"
Public Const SET_CONTRACT_NO As String = "ContractNo"
Public Const SET_PO As String = "PO"
Public Const SET_ORDER_NO As String = "OrderNo"
Public Const SET_HSCN_PATH As String = "HSCNPath"
Public Const SET_COMPANY_PATH As String = "CompanyPath"
Public Const SET_ORDER_PATH As String = "OrderPath"
Public Const SET_OUTPUT_PATH As String = "OutputPath"
Public Const SET_PRICE_FILE As String = "PriceFilePath"
Public Const SET_PRICE_SHEET As String = "PriceSheetName"
Public Const SET_PRICE_HEADER_ROW As String = "PriceHeaderRow"
Public Const SET_PRICE_NAME_COL As String = "PriceNameCol"
Public Const SET_PRICE_END_COL As String = "PriceEndCol"
Public Const SET_PRICE_UNIT_COL As String = "PriceUnitCol"

' --- 業務形式 ---
Public Const SERVICE_MODE_TM As String = "T&M"
Public Const SERVICE_MODE_LUMPSUM As String = "一括請負"

' --- 集計データ構造の列インデックス（内部データ格納用） ---
' 集計結果配列の列定義
Public Const COL_PERSON_NAME As Long = 0      ' 氏名
Public Const COL_LEVEL As Long = 1            ' 等級
Public Const COL_WORK_PERIOD As Long = 2      ' 作業期間
Public Const COL_WORK_CONTENT As Long = 3     ' 作業内容
Public Const COL_WORK_NO As Long = 4          ' 作業番号
Public Const COL_TOTAL_HOURS As Long = 5      ' 総工時
Public Const COL_MAN_MONTH As Long = 6        ' 人月
Public Const COL_UNIT_PRICE As Long = 7       ' 契約単価
Public Const COL_COMPANY As Long = 8          ' 会社名
Public Const COL_DAILY_DATA As Long = 9       ' 日別データ（Dictionary）
```

---

## Step 10：標準モジュール modUtils のコード

**このステップでやること**：各種ユーティリティ関数（簡繁体変換、時間計算など）をまとめた標準モジュールを作成します。

### 10-1. 標準モジュールの挿入

1. VBEでメニュー「挿入」→「標準モジュール」をクリック
2. プロパティウィンドウで **(オブジェクト名)** を **modUtils** に変更します

### 10-2. modUtils のコード

以下のコードを**すべてコピー**して、modUtilsのコード編集エリアに貼り付けてください。

```vba
Option Explicit

'==================================================
' ユーティリティ関数モジュール
' 簡繁体変換、時間計算、設定読み書き、ファイル操作などの
' 汎用関数を集約しています
'==================================================

'--------------------------------------------------
' 簡繁体変換
'--------------------------------------------------

' 簡繁体変換辞書（メモリキャッシュ用）
Private m_dictSimplToTrad As Object

''' 簡繁体変換辞書を简繁对照シートから読み込む
Private Function GetSimplTradDict() As Object
    If Not m_dictSimplToTrad Is Nothing Then
        Set GetSimplTradDict = m_dictSimplToTrad
        Exit Function
    End If
    
    Set m_dictSimplToTrad = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_SIMP_TRAD)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set GetSimplTradDict = m_dictSimplToTrad
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow  ' 1行目は見出し想定
        Dim simp As String, trad As String
        simp = Trim(CStr(ws.Cells(i, 1).Value))
        trad = Trim(CStr(ws.Cells(i, 2).Value))
        If Len(simp) > 0 And Len(trad) > 0 Then
            If Not m_dictSimplToTrad.Exists(simp) Then
                m_dictSimplToTrad.Add simp, trad
            End If
        End If
    Next i
    
    Set GetSimplTradDict = m_dictSimplToTrad
End Function

''' 簡体字を含む文字列を繁体字に正規化する
''' 簡繁对照シートのマッピングに従って1文字ずつ変換します
Public Function NormalizeName(ByVal srcName As String) As String
    If Len(srcName) = 0 Then
        NormalizeName = ""
        Exit Function
    End If
    
    Dim dict As Object
    Set dict = GetSimplTradDict()
    
    Dim result As String
    result = ""
    
    Dim i As Long
    For i = 1 To Len(srcName)
        Dim ch As String
        ch = Mid(srcName, i, 1)
        If dict.Exists(ch) Then
            result = result & dict(ch)
        Else
            result = result & ch
        End If
    Next i
    
    NormalizeName = result
End Function

'--------------------------------------------------
' 勤務時間計算
'--------------------------------------------------

''' 時刻を0.5時間単位で切り捨てる（下方丸め）
''' 例：9:16 → 9:00、18:29 → 18:00
Public Function RoundDownTime(ByVal dt As Date) As Double
    Dim hours As Double
    hours = Hour(dt) + Minute(dt) / 60#
    
    ' 0.5時間単位で切り捨て
    RoundDownTime = Int(hours / ROUND_UNIT_HOURS) * ROUND_UNIT_HOURS
End Function

''' 値を0.5時間単位で切り上げる（上方丸め）
Public Function RoundUpTime(ByVal hours As Double) As Double
    RoundUpTime = Application.WorksheetFunction.Ceiling(hours, ROUND_UNIT_HOURS)
End Function

''' 出勤時刻と退勤時刻から勤務時間を計算する
''' 昼休憩（1時間）を差し引きます
Public Function CalcWorkHours(ByVal timeIn As Date, ByVal timeOut As Date) As Double
    Dim startHours As Double, endHours As Double
    startHours = RoundDownTime(timeIn)
    endHours = RoundDownTime(timeOut)
    
    Dim rawHours As Double
    rawHours = endHours - startHours
    
    ' 昼休憩（1時間）を差し引く
    ' ただし勤務が4時間未満の場合は昼休憩を差し引かない
    If rawHours >= 4# Then
        rawHours = rawHours - LUNCH_BREAK_HOURS
    End If
    
    If rawHours < 0 Then rawHours = 0
    
    CalcWorkHours = rawHours
End Function

''' 勤務時間から残業時間を計算する
''' 標準勤務時間（8h）を超えた分を0.5時間単位で切り上げ
Public Function CalcOvertimeHours(ByVal workHours As Double) As Double
    If workHours <= STANDARD_WORK_HOURS Then
        CalcOvertimeHours = 0
    Else
        CalcOvertimeHours = RoundUpTime(workHours - STANDARD_WORK_HOURS)
    End If
End Function

''' 総勤務時間から人月を計算する
Public Function CalcManMonth(ByVal totalHours As Double) As Double
    If HOURS_PER_MAN_MONTH = 0 Then
        CalcManMonth = 0
    Else
        CalcManMonth = totalHours / HOURS_PER_MAN_MONTH
    End If
End Function

'--------------------------------------------------
' 設定（Settingsシート）読み書き
'--------------------------------------------------

''' Settingsシートから指定キーの値を取得する
Public Function GetSetting(ByVal key As String) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(Trim(CStr(ws.Cells(i, 1).Value)), key, vbTextCompare) = 0 Then
            GetSetting = Trim(CStr(ws.Cells(i, 2).Value))
            Exit Function
        End If
    Next i
    
    GetSetting = ""
End Function

''' Settingsシートに指定キーの値を保存する
Public Sub SaveSetting(ByVal key As String, ByVal value As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(Trim(CStr(ws.Cells(i, 1).Value)), key, vbTextCompare) = 0 Then
            ws.Cells(i, 2).Value = value
            Exit Sub
        End If
    Next i
    
    ' キーが見つからない場合、末尾に追加
    ws.Cells(lastRow + 1, 1).Value = key
    ws.Cells(lastRow + 1, 2).Value = value
End Sub

'--------------------------------------------------
' ファイル/フォルダ操作
'--------------------------------------------------

''' ファイル選択ダイアログを表示し、選択されたファイルパスを返す
Public Function BrowseFile(Optional ByVal title As String = "ファイルを選択", _
                           Optional ByVal filter As String = "Excelファイル (*.xlsx;*.xlsm),*.xlsx;*.xlsm") As String
    Dim fd As Object
    Set fd = Application.FileDialog(3) ' msoFileDialogFilePicker
    
    With fd
        .Title = title
        .Filters.Clear
        .Filters.Add "Excelファイル", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            BrowseFile = .SelectedItems(1)
        Else
            BrowseFile = ""
        End If
    End With
End Function

''' フォルダ選択ダイアログを表示し、選択されたフォルダパスを返す
Public Function BrowseFolder(Optional ByVal title As String = "フォルダを選択") As String
    Dim fd As Object
    Set fd = Application.FileDialog(4) ' msoFileDialogFolderPicker
    
    With fd
        .Title = title
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            BrowseFolder = .SelectedItems(1)
        Else
            BrowseFolder = ""
        End If
    End With
End Function

''' 指定フォルダ内のExcelファイル一覧を取得する
''' 戻り値：ファイルパスの配列
Public Function GetExcelFiles(ByVal folderPath As String) As Variant
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        GetExcelFiles = Array()
        Exit Function
    End If
    
    Dim fileList As Object
    Set fileList = CreateObject("System.Collections.ArrayList")
    
    Dim fileName As String
    fileName = Dir(folderPath & "\*.xlsx")
    Do While fileName <> ""
        fileList.Add folderPath & "\" & fileName
        fileName = Dir()
    Loop
    
    fileName = Dir(folderPath & "\*.xls")
    Do While fileName <> ""
        fileList.Add folderPath & "\" & fileName
        fileName = Dir()
    Loop
    
    If fileList.Count = 0 Then
        GetExcelFiles = Array()
    Else
        GetExcelFiles = fileList.ToArray()
    End If
End Function

'--------------------------------------------------
' その他ユーティリティ
'--------------------------------------------------

''' 現在時刻の文字列表現を返す（ログ出力用）
Public Function NowString() As String
    NowString = Format(Now, "yyyy/mm/dd HH:MM:SS")
End Function

''' 列番号（1-based）を列文字（A, B, ..., Z, AA, ...）に変換
Public Function ColNumToLetter(ByVal colNum As Long) As String
    ColNumToLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

''' 列文字（A〜Z）を列番号（1-based）に変換
Public Function ColLetterToNum(ByVal colLetter As String) As Long
    ColLetterToNum = Range(colLetter & "1").Column
End Function
```

---

## Step 11：標準モジュール modMain のコード

**このステップでやること**：メインの処理ロジック（勤怠集計、ファイル生成）をまとめた標準モジュールを作成します。

### 11-1. 標準モジュールの挿入

1. VBEでメニュー「挿入」→「標準モジュール」をクリック
2. プロパティウィンドウで **(オブジェクト名)** を **modMain** に変更します

### 11-2. modMain のコード

以下のコードを**すべてコピー**して、modMainのコード編集エリアに貼り付けてください。

```vba
Option Explicit

'==================================================
' メイン処理モジュール
' 勤怠集計のコアロジックと請求書生成処理
'==================================================

' 集計結果を保持するモジュールレベル変数
Private m_summaryData As Object  ' Dictionary: Key=正規化氏名, Value=集計結果配列

'--------------------------------------------------
' 公開プロパティ：集計結果
'--------------------------------------------------

''' 集計結果を返す
Public Property Get SummaryData() As Object
    Set SummaryData = m_summaryData
End Property

''' 集計結果があるか
Public Property Get HasSummaryData() As Boolean
    HasSummaryData = Not (m_summaryData Is Nothing)
End Property

'--------------------------------------------------
' 勤怠集計 メイン処理
'--------------------------------------------------

''' 選択された勤怠ファイルを集計する
''' filePaths：集計対象のExcelファイルパス配列
''' location：拠点名（北京/大連/河南）
''' periodFilter：業務期間フィルター（空文字の場合は全期間）
''' 戻り値：集計成功なら True
Public Function ExecuteAggregation(ByVal filePaths As Variant, _
                                    ByVal location As String, _
                                    ByVal periodFilter As String) As Boolean
    On Error GoTo ErrHandler
    
    ' 集計データの初期化
    Set m_summaryData = CreateObject("Scripting.Dictionary")
    
    ' 単価表を読み込む
    Dim priceDict As Object
    Set priceDict = LoadPriceTable()
    
    ' 各ファイルを処理
    Dim i As Long
    For i = LBound(filePaths) To UBound(filePaths)
        Dim filePath As String
        filePath = CStr(filePaths(i))
        
        ' 期間フィルターがあれば、ファイル名で簡易フィルタリング
        If Len(periodFilter) > 0 Then
            If Not IsFileInPeriod(filePath, periodFilter) Then
                GoTo NextFile
            End If
        End If
        
        ' 勤怠ファイルを読み取り
        Call ProcessAttendanceFile(filePath, location, priceDict)
        
NextFile:
    Next i
    
    ExecuteAggregation = True
    Exit Function
    
ErrHandler:
    Debug.Print "ExecuteAggregation Error: " & Err.Description
    ExecuteAggregation = False
End Function

''' 単価表を読み込み、正規化氏名 → 単価情報 の辞書を返す
Private Function LoadPriceTable() As Object
    Dim result As Object
    Set result = CreateObject("Scripting.Dictionary")
    
    Dim filePath As String, sheetName As String, headerRow As Long
    Dim nameCol As Long, endCol As Long, unitCol As Long
    
    filePath = GetSetting(SET_PRICE_FILE)
    sheetName = GetSetting(SET_PRICE_SHEET)
    headerRow = CLng(Val(GetSetting(SET_PRICE_HEADER_ROW)))
    nameCol = ColLetterToNum(GetSetting(SET_PRICE_NAME_COL))
    endCol = ColLetterToNum(GetSetting(SET_PRICE_END_COL))
    unitCol = ColLetterToNum(GetSetting(SET_PRICE_UNIT_COL))
    
    If Len(filePath) = 0 Then
        Set LoadPriceTable = result
        Exit Function
    End If
    
    If Len(Dir(filePath)) = 0 Then
        Set LoadPriceTable = result
        Exit Function
    End If
    
    ' 単価表を読み取り専用で開く
    Dim wb As Workbook
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    If ws Is Nothing Then Set ws = wb.Worksheets(1)
    On Error GoTo 0
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, nameCol).End(xlUp).Row
    
    Dim r As Long
    For r = headerRow + 1 To lastRow
        Dim rawName As String
        rawName = Trim(CStr(ws.Cells(r, nameCol).Value))
        If Len(rawName) > 0 Then
            Dim normName As String
            normName = NormalizeName(rawName)
            
            Dim priceInfo(0 To 2) As Variant
            priceInfo(0) = rawName  ' 元の氏名
            priceInfo(1) = ws.Cells(r, unitCol).Value  ' 契約単価
            priceInfo(2) = ws.Cells(r, endCol).Value   ' 退社日
            
            If Not result.Exists(normName) Then
                result.Add normName, priceInfo
            End If
        End If
    Next r
    
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    
    Set LoadPriceTable = result
End Function

''' 1つの勤怠ファイルを処理する
Private Sub ProcessAttendanceFile(ByVal filePath As String, _
                                   ByVal location As String, _
                                   ByVal priceDict As Object)
    Dim wb As Workbook
    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    
    ' 拠点別にフォーマットが異なるため、拠点に応じた読み取りを行う
    Select Case location
        Case LOCATION_BEIJING
            Call ParseBeijingFormat(ws, priceDict)
        Case LOCATION_DALIAN
            Call ParseDalianFormat(ws, priceDict)
        Case LOCATION_HENAN
            Call ParseHenanFormat(ws, priceDict)
        Case Else
            ' デフォルト：北京形式として処理
            Call ParseBeijingFormat(ws, priceDict)
    End Select
    
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------
' 拠点別フォーマット解析
'--------------------------------------------------

''' 北京拠点の勤怠フォーマットを解析
''' 想定レイアウト：
'''   Row 1〜3：ヘッダー（会社名など）
'''   Row 4〜：氏名行 → その下に日別データ
Private Sub ParseBeijingFormat(ByVal ws As Worksheet, ByVal priceDict As Object)
    ' 簡易実装：全セルをスキャンして氏名行と日付行を検出
    
    Dim companyName As String
    companyName = Trim(CStr(ws.Cells(1, 1).Value))
    
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim r As Long
    For r = 1 To lastRow
        Dim firstCell As String
        firstCell = Trim(CStr(ws.Cells(r, 1).Value))
        
        ' 氏名行の検出（1列目に漢字2〜4文字の氏名）
        If IsNameCell(firstCell) Then
            Dim rawPersonName As String
            rawPersonName = firstCell
            Dim normPersonName As String
            normPersonName = NormalizeName(rawPersonName)
            
            ' 日別データの読み取り
            Dim dailyDict As Object
            Set dailyDict = CreateObject("Scripting.Dictionary")
            
            Dim totalHours As Double
            totalHours = 0
            
            ' r+1行目以降を日別データとして読み取り
            Dim dr As Long
            dr = r + 1
            Do While dr <= lastRow
                Dim dayCell As String
                dayCell = Trim(CStr(ws.Cells(dr, 1).Value))
                
                ' 次の氏名行に当たったら終了
                If IsNameCell(dayCell) Then Exit Do
                
                ' 日付データがあれば処理
                If IsNumeric(dayCell) And Len(dayCell) > 0 Then
                    Dim dayNum As Long
                    dayNum = CLng(dayCell)
                    
                    Dim timeInVal As Variant, timeOutVal As Variant
                    timeInVal = ws.Cells(dr, 2).Value  ' 出勤時刻（2列目想定）
                    timeOutVal = ws.Cells(dr, 3).Value ' 退勤時刻（3列目想定）
                    
                    If IsDate(timeInVal) And IsDate(timeOutVal) Then
                        Dim workHrs As Double, overtime As Double
                        workHrs = CalcWorkHours(CDate(timeInVal), CDate(timeOutVal))
                        overtime = CalcOvertimeHours(workHrs)
                        
                        Dim dayInfo(0 To 3) As Variant
                        dayInfo(0) = dayNum
                        dayInfo(1) = Format(timeInVal, "hh:mm")
                        dayInfo(2) = Format(timeOutVal, "hh:mm")
                        dayInfo(3) = workHrs
                        
                        dailyDict.Add CStr(dr), dayInfo
                        totalHours = totalHours + workHrs
                    End If
                End If
                
                dr = dr + 1
            Loop
            
            Dim manMonth As Double
            manMonth = CalcManMonth(totalHours)
            
            ' 単価マッチング
            Dim unitPrice As Variant
            unitPrice = 0
            If priceDict.Exists(normPersonName) Then
                Dim pInfo As Variant
                pInfo = priceDict(normPersonName)
                unitPrice = pInfo(1)
            End If
            
            ' 集計データを保存
            Dim summary(0 To 9) As Variant
            summary(COL_PERSON_NAME) = rawPersonName
            summary(COL_LEVEL) = ""
            summary(COL_WORK_PERIOD) = ""
            summary(COL_WORK_CONTENT) = ""
            summary(COL_WORK_NO) = ""
            summary(COL_TOTAL_HOURS) = totalHours
            summary(COL_MAN_MONTH) = manMonth
            summary(COL_UNIT_PRICE) = unitPrice
            summary(COL_COMPANY) = companyName
            Set summary(COL_DAILY_DATA) = dailyDict
            
            If m_summaryData.Exists(normPersonName) Then
                ' 既存データに加算
                Dim existing As Variant
                existing = m_summaryData(normPersonName)
                existing(COL_TOTAL_HOURS) = CDbl(existing(COL_TOTAL_HOURS)) + totalHours
                existing(COL_MAN_MONTH) = CalcManMonth(CDbl(existing(COL_TOTAL_HOURS)))
                m_summaryData(normPersonName) = existing
            Else
                m_summaryData.Add normPersonName, summary
            End If
        End If
    Next r
End Sub

''' 大連拠点の勤怠フォーマットを解析（北京と同じ基本構造を想定）
Private Sub ParseDalianFormat(ByVal ws As Worksheet, ByVal priceDict As Object)
    ' 大連形式：北京と同様の構造を想定
    Call ParseBeijingFormat(ws, priceDict)
End Sub

''' 河南拠点の勤怠フォーマットを解析（北京と同じ基本構造を想定）
Private Sub ParseHenanFormat(ByVal ws As Worksheet, ByVal priceDict As Object)
    ' 河南形式：北京と同様の構造を想定
    Call ParseBeijingFormat(ws, priceDict)
End Sub

''' セルの値が氏名かどうかを判定する
''' 簡易判定：漢字2〜4文字で数字を含まない
Private Function IsNameCell(ByVal cellValue As String) As Boolean
    If Len(cellValue) < 2 Or Len(cellValue) > 4 Then
        IsNameCell = False
        Exit Function
    End If
    
    ' 数字を含む場合は氏名ではない
    Dim i As Long
    For i = 1 To Len(cellValue)
        Dim ch As String
        ch = Mid(cellValue, i, 1)
        If ch Like "[0-9０-９]" Then
            IsNameCell = False
            Exit Function
        End If
    Next i
    
    IsNameCell = True
End Function

''' ファイル名から期間フィルターに合致するか簡易判定
Private Function IsFileInPeriod(ByVal filePath As String, ByVal period As String) As Boolean
    Dim fileName As String
    fileName = Dir(filePath)
    
    ' 期間文字列（例：2026/4）がファイル名に含まれるか
    ' 例：period = "2026/4" → ファイル名に "202604" または "2026_04" など
    Dim periodVariants(0 To 3) As String
    periodVariants(0) = Replace(period, "/", "")
    periodVariants(1) = Replace(period, "/", "_")
    periodVariants(2) = Replace(period, "/", "年") & "月"
    periodVariants(3) = period
    
    Dim i As Long
    For i = 0 To 3
        If InStr(1, fileName, periodVariants(i), vbTextCompare) > 0 Then
            IsFileInPeriod = True
            Exit Function
        End If
    Next i
    
    IsFileInPeriod = False
End Function

'--------------------------------------------------
' 請求書生成 メイン処理
'--------------------------------------------------

''' 3種類の請求書を一括生成する
Public Function GenerateAllReports() As Boolean
    On Error GoTo ErrHandler
    
    If Not HasSummaryData Then
        MsgBox "先に勤怠集計を実行してください。", vbExclamation, "確認"
        GenerateAllReports = False
        Exit Function
    End If
    
    Dim outputPath As String
    outputPath = GetSetting(SET_OUTPUT_PATH)
    If Len(outputPath) = 0 Then
        outputPath = ThisWorkbook.Path & "\Output"
    End If
    
    ' 出力フォルダがなければ作成
    If Len(Dir(outputPath, vbDirectory)) = 0 Then
        MkDir outputPath
    End If
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 1. HSCN验收书 生成
    If Not GenerateHSCN(outputPath) Then
        Debug.Print "HSCN验收书 生成に失敗しました"
    End If
    
    ' 2. 会社明細 生成
    If Not GenerateCompanyReport(outputPath) Then
        Debug.Print "会社明細 生成に失敗しました"
    End If
    
    ' 3. 实际报价兼发注书 生成
    If Not GenerateOrderReport(outputPath) Then
        Debug.Print "实际报价兼发注书 生成に失敗しました"
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    GenerateAllReports = True
    Exit Function
    
ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Debug.Print "GenerateAllReports Error: " & Err.Description
    GenerateAllReports = False
End Function

''' HSCN验收书を生成する
Private Function GenerateHSCN(ByVal outputPath As String) As Boolean
    Dim templatePath As String
    templatePath = GetSetting(SET_HSCN_PATH)
    
    If Len(templatePath) = 0 Or Len(Dir(templatePath)) = 0 Then
        MsgBox "HSCN验收书のテンプレートが見つかりません。", vbExclamation
        GenerateHSCN = False
        Exit Function
    End If
    
    ' 出力ファイル名
    Dim outputFile As String
    outputFile = outputPath & "\HSCN验收书_" & Format(Now, "yyyymmdd") & ".xlsx"
    
    ' テンプレートをコピー
    FileCopy templatePath, outputFile
    
    ' 出力ファイルを開いて編集
    Dim wb As Workbook
    Set wb = Workbooks.Open(outputFile)
    
    ' ---- Timesheet シートの生成 ----
    Dim templateWS As Worksheet
    On Error Resume Next
    Set templateWS = wb.Worksheets(TEMPLATE_TIMESHEET)
    On Error GoTo 0
    
    If Not templateWS Is Nothing Then
        ' 各人員のTimesheetシートを作成
        Dim key As Variant
        For Each key In m_summaryData.Keys
            Dim summary As Variant
            summary = m_summaryData(key)
            
            ' Timesheet_Templateをコピー
            templateWS.Copy After:=wb.Worksheets(wb.Worksheets.Count)
            Dim newWS As Worksheet
            Set newWS = wb.Worksheets(wb.Worksheets.Count)
            
            Dim personName As String
            personName = CStr(summary(COL_PERSON_NAME))
            
            ' シート名設定（31文字制限に注意）
            Dim sheetName As String
            sheetName = "Timesheet_" & personName
            If Len(sheetName) > 31 Then sheetName = Left(sheetName, 31)
            On Error Resume Next
            newWS.Name = sheetName
            On Error GoTo 0
            
            ' Row 4：会社名
            newWS.Cells(4, 2).Value = CStr(summary(COL_COMPANY))
            
            ' Row 6：氏名、等級
            newWS.Cells(6, 2).Value = personName
            newWS.Cells(6, 4).Value = CStr(summary(COL_LEVEL))
            
            ' Row 8：年月、総工時、人月
            newWS.Cells(8, 2).Value = Format(Now, "yyyy/mm")
            newWS.Cells(8, 4).Value = CDbl(summary(COL_TOTAL_HOURS))
            newWS.Cells(8, 6).Value = CDbl(summary(COL_MAN_MONTH))
            
            ' Row 11以降：日別データ
            Dim dailyDict As Object
            Set dailyDict = summary(COL_DAILY_DATA)
            
            If Not dailyDict Is Nothing And dailyDict.Count > 0 Then
                Dim rowIdx As Long
                rowIdx = 11
                
                Dim dk As Variant
                For Each dk In dailyDict.Keys
                    Dim dayInfo As Variant
                    dayInfo = dailyDict(dk)
                    
                    newWS.Cells(rowIdx, 1).Value = Month(Now)  ' 月
                    newWS.Cells(rowIdx, 2).Value = dayInfo(0)   ' 日
                    newWS.Cells(rowIdx, 3).Value = ""           ' 作業内容（後で手動入力）
                    newWS.Cells(rowIdx, 4).Value = dayInfo(1)   ' 上班时间
                    newWS.Cells(rowIdx, 5).Value = dayInfo(2)   ' 下班时间
                    newWS.Cells(rowIdx, 6).Value = 0            ' 加班（テンプレの計算式想定）
                    newWS.Cells(rowIdx, 7).Value = 0            ' 缺勤
                    newWS.Cells(rowIdx, 8).Value = dayInfo(3)   ' 工时
                    
                    rowIdx = rowIdx + 1
                Next dk
            End If
        Next key
        
        ' Timesheet_Templateを非表示に
        templateWS.Visible = xlSheetHidden
    End If
    
    ' ---- Acceptance Requestシートの編集 ----
    Dim accWS As Worksheet
    On Error Resume Next
    Set accWS = wb.Worksheets(TEMPLATE_ACCEPTANCE)
    On Error GoTo 0
    
    If Not accWS Is Nothing Then
        ' Row 5〜12：プロジェクト設定から埋め込み
        accWS.Cells(5, 2).Value = GetSetting(SET_PROJECT_NAME)
        accWS.Cells(6, 2).Value = GetSetting(SET_SUPPLIER)
        accWS.Cells(7, 2).Value = GetSetting(SET_SERVICE_MODE)
        accWS.Cells(8, 2).Value = GetSetting(SET_SERVICE_CONTENT)
        accWS.Cells(9, 2).Value = GetSetting(SET_DELIVERABLE)
        accWS.Cells(10, 2).Value = GetSetting(SET_MILESTONE)
        accWS.Cells(11, 2).Value = GetSetting(SET_CONTRACT_NO)
        accWS.Cells(12, 2).Value = GetSetting(SET_PO)
        
        ' Row 16以降：人員別集計データ
        Dim accRow As Long
        accRow = 16
        Dim idx As Long
        idx = 1
        
        For Each key In m_summaryData.Keys
            summary = m_summaryData(key)
            
            accWS.Cells(accRow, 1).Value = idx
            accWS.Cells(accRow, 2).Value = CStr(summary(COL_PERSON_NAME))
            accWS.Cells(accRow, 3).Value = CStr(summary(COL_WORK_PERIOD))
            accWS.Cells(accRow, 4).Value = CStr(summary(COL_WORK_CONTENT))
            accWS.Cells(accRow, 5).Value = CDbl(summary(COL_TOTAL_HOURS))
            accWS.Cells(accRow, 6).Value = CDbl(summary(COL_MAN_MONTH))
            accWS.Cells(accRow, 7).Value = CStr(summary(COL_LEVEL))
            accWS.Cells(accRow, 8).Value = CDbl(summary(COL_UNIT_PRICE))
            ' 金額列はテンプレートの計算式に任せる
            
            accRow = accRow + 1
            idx = idx + 1
        Next key
    End If
    
    wb.Close SaveChanges:=True
    GenerateHSCN = True
End Function

''' 会社明細を生成する
Private Function GenerateCompanyReport(ByVal outputPath As String) As Boolean
    Dim templatePath As String
    templatePath = GetSetting(SET_COMPANY_PATH)
    
    If Len(templatePath) = 0 Or Len(Dir(templatePath)) = 0 Then
        MsgBox "会社明細のテンプレートが見つかりません。", vbExclamation
        GenerateCompanyReport = False
        Exit Function
    End If
    
    Dim outputFile As String
    outputFile = outputPath & "\会社明細_" & Format(Now, "yyyymmdd") & ".xlsx"
    
    FileCopy templatePath, outputFile
    
    Dim wb As Workbook
    Set wb = Workbooks.Open(outputFile)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    
    ' Row 2以降に人員別データを埋め込み
    Dim dataRow As Long
    dataRow = 2
    Dim idx As Long
    idx = 1
    
    Dim key As Variant
    For Each key In m_summaryData.Keys
        Dim summary As Variant
        summary = m_summaryData(key)
        
        ws.Cells(dataRow, 1).Value = idx
        ws.Cells(dataRow, 2).Value = CStr(summary(COL_PERSON_NAME))
        ws.Cells(dataRow, 3).Value = CStr(summary(COL_WORK_PERIOD))
        ws.Cells(dataRow, 4).Value = CStr(summary(COL_WORK_CONTENT))
        ws.Cells(dataRow, 5).Value = CDbl(summary(COL_TOTAL_HOURS))
        ws.Cells(dataRow, 6).Value = CDbl(summary(COL_MAN_MONTH))
        ws.Cells(dataRow, 7).Value = CStr(summary(COL_LEVEL))
        ws.Cells(dataRow, 8).Value = CDbl(summary(COL_UNIT_PRICE))
        
        dataRow = dataRow + 1
        idx = idx + 1
    Next key
    
    wb.Close SaveChanges:=True
    GenerateCompanyReport = True
End Function

''' 实际报价兼发注书を生成する
Private Function GenerateOrderReport(ByVal outputPath As String) As Boolean
    Dim templatePath As String
    templatePath = GetSetting(SET_ORDER_PATH)
    
    If Len(templatePath) = 0 Or Len(Dir(templatePath)) = 0 Then
        MsgBox "实际报价兼发注书のテンプレートが見つかりません。", vbExclamation
        GenerateOrderReport = False
        Exit Function
    End If
    
    Dim outputFile As String
    outputFile = outputPath & "\实际报价兼发注书_" & Format(Now, "yyyymmdd") & ".xlsx"
    
    FileCopy templatePath, outputFile
    
    Dim wb As Workbook
    Set wb = Workbooks.Open(outputFile)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    
    ' Row 14以降（表一）に人員別データを埋め込み
    Dim dataRow As Long
    dataRow = 14
    Dim idx As Long
    idx = 1
    
    Dim key As Variant
    For Each key In m_summaryData.Keys
        Dim summary As Variant
        summary = m_summaryData(key)
        
        ws.Cells(dataRow, 1).Value = idx
        ws.Cells(dataRow, 2).Value = CStr(summary(COL_PERSON_NAME))
        ws.Cells(dataRow, 3).Value = CStr(summary(COL_WORK_PERIOD))
        ws.Cells(dataRow, 4).Value = CStr(summary(COL_WORK_CONTENT))
        ws.Cells(dataRow, 5).Value = CDbl(summary(COL_TOTAL_HOURS))
        ws.Cells(dataRow, 6).Value = CDbl(summary(COL_MAN_MONTH))
        ws.Cells(dataRow, 7).Value = CStr(summary(COL_LEVEL))
        ws.Cells(dataRow, 8).Value = CDbl(summary(COL_UNIT_PRICE))
        
        dataRow = dataRow + 1
        idx = idx + 1
    Next key
    
    wb.Close SaveChanges:=True
    GenerateOrderReport = True
End Function
```

---
## Step 12：frmMain のコード

**このステップでやること**：メイン画面（frmMain）のVBAコードを記述します。

### 操作手順

1. VBEのプロジェクトエクスプローラーで `frmMain` をダブルクリック
2. 表示されたコード編集エリアに以下のコードを**すべてコピー＆ペースト**します

```vba
Option Explicit

'==================================================
' frmMain コード
' メイン画面のイベント処理
'==================================================

'--------------------------------------------------
' フォーム初期化
'--------------------------------------------------

Private Sub UserForm_Initialize()
    ' 拠点コンボボックスに選択肢を追加
    cboLocation.AddItem LOCATION_BEIJING
    cboLocation.AddItem LOCATION_DALIAN
    cboLocation.AddItem LOCATION_HENAN
    
    ' 前回の設定を復元
    Dim lastLocation As String
    lastLocation = GetSetting("LastLocation")
    If Len(lastLocation) > 0 Then
        Dim i As Long
        For i = 0 To cboLocation.ListCount - 1
            If cboLocation.List(i) = lastLocation Then
                cboLocation.ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboLocation.ListIndex = 0
    End If
    
    ' フォルダパスを復元
    txtFolderPath.Text = GetSetting("LastFolderPath")
    
    ' 担当者コンボボックスに選択肢を追加
    cboPerson.AddItem "(すべて)"
    cboPerson.ListIndex = 0
    
    ' ステータス初期表示
    lblStatus.Caption = "待機中"
    lblCount.Caption = "選択件数：0 件"
    
    ' ファイル一覧を更新
    If Len(txtFolderPath.Text) > 0 Then
        Call RefreshFileList
    End If
End Sub

'--------------------------------------------------
' フォルダパス変更時
'--------------------------------------------------

Private Sub txtFolderPath_Change()
    Call RefreshFileList
End Sub

'--------------------------------------------------
' 参照ボタンクリック
'--------------------------------------------------

Private Sub btnBrowse_Click()
    Dim folderPath As String
    folderPath = BrowseFolder("勤怠表フォルダを選択")
    
    If Len(folderPath) > 0 Then
        txtFolderPath.Text = folderPath
        SaveSetting "LastFolderPath", folderPath
    End If
End Sub

'--------------------------------------------------
' 拠点選択変更時
'--------------------------------------------------

Private Sub cboLocation_Change()
    If cboLocation.ListIndex >= 0 Then
        SaveSetting "LastLocation", cboLocation.Text
        ' 拠点が変わったら担当者一覧を更新
        Call RefreshPersonList
        Call RefreshFileList
    End If
End Sub

'--------------------------------------------------
' 担当者選択変更時
'--------------------------------------------------

Private Sub cboPerson_Change()
    Call RefreshFileList
End Sub

'--------------------------------------------------
' 業務期間変更時
'--------------------------------------------------

Private Sub txtPeriod_Change()
    ' 期間フィルターとして使用（ファイル名に期間文字列が含まれるか）
End Sub

'--------------------------------------------------
' 全选チェックボックス
'--------------------------------------------------

Private Sub chkSelectAll_Click()
    Dim i As Long
    For i = 0 To lstFiles.ListCount - 1
        lstFiles.Selected(i) = chkSelectAll.Value
    Next i
    Call UpdateCount
End Sub

'--------------------------------------------------
' ファイルリスト選択変更時
'--------------------------------------------------

Private Sub lstFiles_Change()
    Call UpdateCount
End Sub

''' 選択件数ラベルを更新
Private Sub UpdateCount()
    Dim count As Long
    count = 0
    
    Dim i As Long
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then
            count = count + 1
        End If
    Next i
    
    lblCount.Caption = "選択件数：" & count & " 件"
End Sub

'--------------------------------------------------
' プロジェクト設定ボタン
'--------------------------------------------------

Private Sub btnProject_Click()
    frmProject.Show vbModal
End Sub

'--------------------------------------------------
' 人員設定ボタン
'--------------------------------------------------

Private Sub btnPerson_Click()
    frmPerson.Show vbModal
End Sub

'--------------------------------------------------
' 単価表設定ボタン
'--------------------------------------------------

Private Sub btnPrice_Click()
    frmPrice.Show vbModal
End Sub

'--------------------------------------------------
' 勤怠集計実行ボタン
'--------------------------------------------------

Private Sub btnExecute_Click()
    ' 選択ファイルの収集
    Dim selectedFiles As Object
    Set selectedFiles = CreateObject("System.Collections.ArrayList")
    
    Dim i As Long
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then
            Dim filePath As String
            filePath = CStr(lstFiles.List(i, 0))
            selectedFiles.Add filePath
        End If
    Next i
    
    If selectedFiles.Count = 0 Then
        MsgBox "集計するファイルを選択してください。", vbExclamation, "確認"
        Exit Sub
    End If
    
    ' 拠点・期間を取得
    Dim location As String
    location = ""
    If cboLocation.ListIndex >= 0 Then
        location = cboLocation.Text
    End If
    
    Dim period As String
    period = Trim(txtPeriod.Text)
    
    ' ステータス更新
    lblStatus.Caption = "勤怠集計中..."
    lblStatus.ForeColor = &H8000&
    DoEvents
    
    ' 集計実行
    Dim startTime As Double
    startTime = Timer
    
    Dim filePaths() As String
    ReDim filePaths(0 To selectedFiles.Count - 1)
    For i = 0 To selectedFiles.Count - 1
        filePaths(i) = CStr(selectedFiles(i))
    Next i
    
    Dim success As Boolean
    success = ExecuteAggregation(filePaths, location, period)
    
    Dim elapsed As Double
    elapsed = Timer - startTime
    
    If success Then
        lblStatus.Caption = "集計完了（" & Format(elapsed, "0.0") & "秒）"
        lblStatus.ForeColor = &H8000&
        
        ' 処理履歴に追加
        Dim logMsg As String
        logMsg = NowString() & " - 集計完了：" & selectedFiles.Count & "ファイル → " & _
                 SummaryData.Count & "名"
        lstHistory.AddItem logMsg, 0
        
        MsgBox "勤怠集計が完了しました。" & vbCrLf & vbCrLf & _
               "集計ファイル数：" & selectedFiles.Count & vbCrLf & _
               "集計人員数：" & SummaryData.Count & "名" & vbCrLf & _
               "処理時間：" & Format(elapsed, "0.0") & "秒", vbInformation, "集計完了"
    Else
        lblStatus.Caption = "集計に失敗しました"
        lblStatus.ForeColor = &HFF&
        MsgBox "勤怠集計中にエラーが発生しました。", vbCritical, "エラー"
    End If
End Sub

'--------------------------------------------------
' 3ファイル一括生成ボタン
'--------------------------------------------------

Private Sub btnGenerate_Click()
    If Not HasSummaryData Then
        MsgBox "先に勤怠集計を実行してください。", vbExclamation, "確認"
        Exit Sub
    End If
    
    ' 出力先の確認
    Dim outPath As String
    outPath = GetSetting(SET_OUTPUT_PATH)
    If Len(outPath) = 0 Then
        outPath = ThisWorkbook.Path & "\Output"
    End If
    
    Dim answer As VbMsgBoxResult
    answer = MsgBox("以下のフォルダに3種類の請求書を出力します。" & vbCrLf & vbCrLf & _
                    outPath & vbCrLf & vbCrLf & _
                    "続行しますか？", vbYesNo + vbQuestion, "出力確認")
    
    If answer <> vbYes Then
        Exit Sub
    End If
    
    ' ステータス更新
    lblStatus.Caption = "請求書生成中..."
    lblStatus.ForeColor = &H8000&
    DoEvents
    
    ' 生成実行
    Dim startTime As Double
    startTime = Timer
    
    Dim success As Boolean
    success = GenerateAllReports()
    
    Dim elapsed As Double
    elapsed = Timer - startTime
    
    If success Then
        lblStatus.Caption = "請求書生成完了（" & Format(elapsed, "0.0") & "秒）"
        lblStatus.ForeColor = &H8000&
        
        ' 処理履歴に追加
        Dim logMsg As String
        logMsg = NowString() & " - 請求書生成完了 → " & outPath
        lstHistory.AddItem logMsg, 0
        
        ' 出力フォルダを開くか確認
        answer = MsgBox("3種類の請求書を生成しました。" & vbCrLf & vbCrLf & _
                        "出力先：" & outPath & vbCrLf & vbCrLf & _
                        "出力フォルダを開きますか？", vbYesNo + vbInformation, "生成完了")
        
        If answer = vbYes Then
            Call OpenOutputFolder
        End If
    Else
        lblStatus.Caption = "生成に失敗しました"
        lblStatus.ForeColor = &HFF&
        MsgBox "請求書生成中にエラーが発生しました。", vbCritical, "エラー"
    End If
End Sub

'--------------------------------------------------
' 出力フォルダを開くボタン
'--------------------------------------------------

Private Sub btnOpenFolder_Click()
    Call OpenOutputFolder
End Sub

''' 出力フォルダをエクスプローラーで開く
Private Sub OpenOutputFolder()
    Dim outPath As String
    outPath = GetSetting(SET_OUTPUT_PATH)
    If Len(outPath) = 0 Then
        outPath = ThisWorkbook.Path & "\Output"
    End If
    
    ' フォルダがなければ作成
    If Len(Dir(outPath, vbDirectory)) = 0 Then
        MkDir outPath
    End If
    
    ' エクスプローラーで開く
    Shell "explorer.exe " & Chr(34) & outPath & Chr(34), vbNormalFocus
End Sub

'--------------------------------------------------
' ファイル一覧更新
'--------------------------------------------------

Private Sub RefreshFileList()
    lstFiles.Clear
    
    Dim folderPath As String
    folderPath = Trim(txtFolderPath.Text)
    
    If Len(folderPath) = 0 Then Exit Sub
    
    Dim files As Variant
    files = GetExcelFiles(folderPath)
    
    ' 担当者フィルター
    Dim filterPerson As String
    filterPerson = ""
    If cboPerson.ListIndex > 0 Then
        filterPerson = cboPerson.Text
    End If
    
    ' 期間フィルター
    Dim filterPeriod As String
    filterPeriod = Trim(txtPeriod.Text)
    
    Dim i As Long
    For i = LBound(files) To UBound(files)
        Dim filePath As String
        filePath = CStr(files(i))
        Dim fileName As String
        fileName = Dir(filePath)
        
        ' 担当者フィルター（ファイル名に担当者名が含まれるか）
        If Len(filterPerson) > 0 Then
            If InStr(1, fileName, filterPerson, vbTextCompare) = 0 Then
                GoTo SkipFile
            End If
        End If
        
        ' 期間フィルター（modMain.IsFileInPeriod を使用）
        If Len(filterPeriod) > 0 Then
            If Not IsFileInPeriod(filePath, filterPeriod) Then
                GoTo SkipFile
            End If
        End If
        
        ' ファイル情報を取得
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        
        Dim fileObj As Object
        Set fileObj = fso.GetFile(filePath)
        
        Dim fileSize As String
        If fileObj.Size < 1024 Then
            fileSize = fileObj.Size & " B"
        ElseIf fileObj.Size < 1048576 Then
            fileSize = Format(fileObj.Size / 1024, "0") & " KB"
        Else
            fileSize = Format(fileObj.Size / 1048576, "0.0") & " MB"
        End If
        
        Dim modifyDate As String
        modifyDate = Format(fileObj.DateLastModified, "yyyy/mm/dd HH:MM")
        
        ' リストに追加（3列）
        lstFiles.AddItem fileName
        Dim rowIdx As Long
        rowIdx = lstFiles.ListCount - 1
        lstFiles.List(rowIdx, 1) = fileSize
        lstFiles.List(rowIdx, 2) = modifyDate
        
SkipFile:
    Next i
    
    ' 全选チェックをリセット
    chkSelectAll.Value = False
    Call UpdateCount
End Sub

'--------------------------------------------------
' 担当者一覧更新
'--------------------------------------------------

Private Sub RefreshPersonList()
    ' 現在の選択を保存
    Dim lastPerson As String
    lastPerson = ""
    If cboPerson.ListIndex >= 0 Then
        lastPerson = cboPerson.Text
    End If
    
    ' コンボボックスをクリアして再構築
    cboPerson.Clear
    cboPerson.AddItem "(すべて)"
    
    ' 拠点フォルダ内のファイルから担当者名を抽出
    ' （実際の運用ではファイルの中身を読んで担当者名を抽出することを推奨）
    
    ' 前回の選択を復元
    If Len(lastPerson) > 0 Then
        Dim i As Long
        For i = 0 To cboPerson.ListCount - 1
            If cboPerson.List(i) = lastPerson Then
                cboPerson.ListIndex = i
                Exit For
            End If
        Next i
    Else
        cboPerson.ListIndex = 0
    End If
End Sub

'--------------------------------------------------
' フォーム終了時
'--------------------------------------------------

Private Sub UserForm_Terminate()
    ' 設定を保存
    If Len(txtFolderPath.Text) > 0 Then
        SaveSetting "LastFolderPath", txtFolderPath.Text
    End If
End Sub
```

---

## Step 13：frmProject のコード

**このステップでやること**：プロジェクト設定画面（frmProject）のVBAコードを記述します。

```vba
Option Explicit

'==================================================
' frmProject コード
'==================================================

Private Sub UserForm_Initialize()
    cboServiceMode.AddItem SERVICE_MODE_TM
    cboServiceMode.AddItem SERVICE_MODE_LUMPSUM
    
    txtProjectName.Text = GetSetting(SET_PROJECT_NAME)
    txtSupplier.Text = GetSetting(SET_SUPPLIER)
    txtServiceContent.Text = GetSetting(SET_SERVICE_CONTENT)
    txtDeliverable.Text = GetSetting(SET_DELIVERABLE)
    txtMilestone.Text = GetSetting(SET_MILESTONE)
    txtContractNo.Text = GetSetting(SET_CONTRACT_NO)
    txtPO.Text = GetSetting(SET_PO)
    txtOrderNo.Text = GetSetting(SET_ORDER_NO)
    
    Dim svcMode As String
    svcMode = GetSetting(SET_SERVICE_MODE)
    If Len(svcMode) > 0 Then cboServiceMode.Text = svcMode Else cboServiceMode.ListIndex = 0
    
    txtHSCNPath.Text = GetSetting(SET_HSCN_PATH)
    txtCompanyPath.Text = GetSetting(SET_COMPANY_PATH)
    txtOrderPath.Text = GetSetting(SET_ORDER_PATH)
    txtOutputPath.Text = GetSetting(SET_OUTPUT_PATH)
End Sub

Private Sub btnHSCNBrowse_Click()
    Dim fp As String: fp = BrowseFile("HSCN验收书テンプレートを選択")
    If Len(fp) > 0 Then txtHSCNPath.Text = fp
End Sub

Private Sub btnCompanyBrowse_Click()
    Dim fp As String: fp = BrowseFile("会社明細テンプレートを選択")
    If Len(fp) > 0 Then txtCompanyPath.Text = fp
End Sub

Private Sub btnOrderBrowse_Click()
    Dim fp As String: fp = BrowseFile("实际报价兼发注书テンプレートを選択")
    If Len(fp) > 0 Then txtOrderPath.Text = fp
End Sub

Private Sub btnOutputBrowse_Click()
    Dim fp As String: fp = BrowseFolder("出力先フォルダを選択")
    If Len(fp) > 0 Then txtOutputPath.Text = fp
End Sub

Private Sub btnSaveProject_Click()
    SaveSetting SET_PROJECT_NAME, txtProjectName.Text
    SaveSetting SET_SUPPLIER, txtSupplier.Text
    SaveSetting SET_SERVICE_MODE, cboServiceMode.Text
    SaveSetting SET_SERVICE_CONTENT, txtServiceContent.Text
    SaveSetting SET_DELIVERABLE, txtDeliverable.Text
    SaveSetting SET_MILESTONE, txtMilestone.Text
    SaveSetting SET_CONTRACT_NO, txtContractNo.Text
    SaveSetting SET_PO, txtPO.Text
    SaveSetting SET_ORDER_NO, txtOrderNo.Text
    SaveSetting SET_HSCN_PATH, txtHSCNPath.Text
    SaveSetting SET_COMPANY_PATH, txtCompanyPath.Text
    SaveSetting SET_ORDER_PATH, txtOrderPath.Text
    SaveSetting SET_OUTPUT_PATH, txtOutputPath.Text
    ThisWorkbook.Save
    Unload Me
End Sub

Private Sub btnCancelProject_Click()
    Unload Me
End Sub
```

---

## Step 14：frmPerson のコード

**このステップでやること**：人員設定画面（frmPerson）のVBAコードを記述します。

```vba
Option Explicit

'==================================================
' frmPerson コード
'==================================================

Private m_editIndex As Long

Private Sub UserForm_Initialize()
    m_editIndex = -1
    cboLevel.AddItem "Senior 1"
    cboLevel.AddItem "Senior 2"
    cboLevel.AddItem "Junior"
    cboLevel.AddItem "Middle"
    cboLevel.ListIndex = 0
    Call LoadPersonList
End Sub

Private Sub LoadPersonList()
    lstPersonList.Clear
    Dim personData As String
    personData = GetSetting("PersonData")
    If Len(personData) = 0 Then Exit Sub
    
    Dim lines() As String: lines = Split(personData, vbCrLf)
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim line As String: line = Trim(lines(i))
        If Len(line) > 0 Then
            Dim cols() As String: cols = Split(line, vbTab)
            If UBound(cols) >= 5 Then
                Dim rowIdx As Long: rowIdx = lstPersonList.ListCount
                lstPersonList.AddItem CStr(rowIdx + 1)
                lstPersonList.List(rowIdx, 1) = cols(0)
                lstPersonList.List(rowIdx, 2) = cols(1)
                lstPersonList.List(rowIdx, 3) = cols(2)
                lstPersonList.List(rowIdx, 4) = cols(3)
                lstPersonList.List(rowIdx, 5) = cols(4)
            End If
        End If
    Next i
End Sub

Private Sub btnAdd_Click()
    If Len(Trim(cboName.Text)) = 0 Then
        MsgBox "氏名を入力してください。", vbExclamation: cboName.SetFocus: Exit Sub
    End If
    Dim rowIdx As Long: rowIdx = lstPersonList.ListCount
    lstPersonList.AddItem CStr(rowIdx + 1)
    lstPersonList.List(rowIdx, 1) = cboName.Text
    lstPersonList.List(rowIdx, 2) = txtWorkPeriod.Text
    lstPersonList.List(rowIdx, 3) = txtWorkContent.Text
    lstPersonList.List(rowIdx, 4) = txtWorkNo.Text
    lstPersonList.List(rowIdx, 5) = cboLevel.Text
    Call ClearInputs
End Sub

Private Sub lstPersonList_Click()
    If lstPersonList.ListIndex < 0 Then Exit Sub
    m_editIndex = lstPersonList.ListIndex
    cboName.Text = lstPersonList.List(m_editIndex, 1)
    txtWorkPeriod.Text = lstPersonList.List(m_editIndex, 2)
    txtWorkContent.Text = lstPersonList.List(m_editIndex, 3)
    txtWorkNo.Text = lstPersonList.List(m_editIndex, 4)
    Dim lvl As String: lvl = lstPersonList.List(m_editIndex, 5)
    Dim i As Long
    For i = 0 To cboLevel.ListCount - 1
        If cboLevel.List(i) = lvl Then cboLevel.ListIndex = i: Exit For
    Next i
End Sub

Private Sub btnUpdate_Click()
    If m_editIndex < 0 Then
        MsgBox "更新する行を選択してください。", vbExclamation: Exit Sub
    End If
    If Len(Trim(cboName.Text)) = 0 Then
        MsgBox "氏名を入力してください。", vbExclamation: Exit Sub
    End If
    lstPersonList.List(m_editIndex, 1) = cboName.Text
    lstPersonList.List(m_editIndex, 2) = txtWorkPeriod.Text
    lstPersonList.List(m_editIndex, 3) = txtWorkContent.Text
    lstPersonList.List(m_editIndex, 4) = txtWorkNo.Text
    lstPersonList.List(m_editIndex, 5) = cboLevel.Text
    Call ClearInputs: m_editIndex = -1
End Sub

Private Sub btnDelete_Click()
    If lstPersonList.ListIndex < 0 Then
        MsgBox "削除する行を選択してください。", vbExclamation: Exit Sub
    End If
    If MsgBox("選択行を削除しますか？", vbYesNo + vbQuestion, "削除確認") <> vbYes Then Exit Sub
    lstPersonList.RemoveItem lstPersonList.ListIndex
    Dim i As Long
    For i = 0 To lstPersonList.ListCount - 1
        lstPersonList.List(i, 0) = CStr(i + 1)
    Next i
    Call ClearInputs: m_editIndex = -1
End Sub

Private Sub btnClearAll_Click()
    If MsgBox("すべての人員データを削除しますか？", vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    lstPersonList.Clear: Call ClearInputs: m_editIndex = -1
End Sub

Private Sub btnSavePerson_Click()
    Dim personData As String: personData = ""
    Dim i As Long
    For i = 0 To lstPersonList.ListCount - 1
        If Len(lstPersonList.List(i, 1)) > 0 Then
            Dim line As String
            line = lstPersonList.List(i, 1) & vbTab & _
                   lstPersonList.List(i, 2) & vbTab & _
                   lstPersonList.List(i, 3) & vbTab & _
                   lstPersonList.List(i, 4) & vbTab & _
                   lstPersonList.List(i, 5)
            If Len(personData) > 0 Then personData = personData & vbCrLf
            personData = personData & line
        End If
    Next i
    SaveSetting "PersonData", personData
    ThisWorkbook.Save: Unload Me
End Sub

Private Sub btnCancelPerson_Click()
    Unload Me
End Sub

Private Sub ClearInputs()
    cboName.Text = "": txtWorkPeriod.Text = ""
    txtWorkContent.Text = "": txtWorkNo.Text = ""
    If cboLevel.ListCount > 0 Then cboLevel.ListIndex = 0
End Sub
```

---

## Step 15：frmPrice のコード

**このステップでやること**：単価表設定画面（frmPrice）のVBAコードを記述します。

```vba
Option Explicit

'==================================================
' frmPrice コード
'==================================================

Private Sub UserForm_Initialize()
    Dim i As Long
    For i = 1 To 26
        Dim col As String: col = Chr(64 + i)
        cboPriceNameCol.AddItem col
        cboPriceEndCol.AddItem col
        cboPriceUnitCol.AddItem col
    Next i
    
    txtPricePath.Text = GetSetting(SET_PRICE_FILE)
    txtPriceHeaderRow.Text = GetSetting(SET_PRICE_HEADER_ROW)
    If Len(txtPriceHeaderRow.Text) = 0 Then txtPriceHeaderRow.Text = "1"
    
    RestoreCombo cboPriceNameCol, GetSetting(SET_PRICE_NAME_COL), "A"
    RestoreCombo cboPriceEndCol, GetSetting(SET_PRICE_END_COL), "B"
    RestoreCombo cboPriceUnitCol, GetSetting(SET_PRICE_UNIT_COL), "C"
    
    If Len(txtPricePath.Text) > 0 And Len(Dir(txtPricePath.Text)) > 0 Then
        Call RefreshSheetList
    End If
    
    RestoreCombo cboPriceSheet, GetSetting(SET_PRICE_SHEET), ""
End Sub

Private Sub RestoreCombo(ByRef cmb As ComboBox, ByVal saved As String, ByVal default As String)
    If Len(saved) > 0 Then
        Dim i As Long
        For i = 0 To cmb.ListCount - 1
            If cmb.List(i) = saved Then cmb.ListIndex = i: Exit Sub
        Next i
    End If
    If Len(default) > 0 Then
        For i = 0 To cmb.ListCount - 1
            If cmb.List(i) = default Then cmb.ListIndex = i: Exit Sub
        Next i
    End If
End Sub

Private Sub btnPriceBrowse_Click()
    Dim fp As String: fp = BrowseFile("単価表ファイルを選択")
    If Len(fp) > 0 Then txtPricePath.Text = fp: Call RefreshSheetList
End Sub

Private Sub RefreshSheetList()
    cboPriceSheet.Clear
    Dim fp As String: fp = Trim(txtPricePath.Text)
    If Len(fp) = 0 Or Len(Dir(fp)) = 0 Then Exit Sub
    
    Application.ScreenUpdating = False
    Dim wb As Workbook: Set wb = Workbooks.Open(fp, ReadOnly:=True)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        cboPriceSheet.AddItem ws.Name
    Next ws
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    If cboPriceSheet.ListCount > 0 Then cboPriceSheet.ListIndex = 0
End Sub

Private Sub btnSavePrice_Click()
    If Not IsNumeric(txtPriceHeaderRow.Text) Then
        MsgBox "見出し行番号には数値を入力してください。", vbExclamation
        txtPriceHeaderRow.SetFocus: Exit Sub
    End If
    SaveSetting SET_PRICE_FILE, txtPricePath.Text
    SaveSetting SET_PRICE_SHEET, cboPriceSheet.Text
    SaveSetting SET_PRICE_HEADER_ROW, txtPriceHeaderRow.Text
    SaveSetting SET_PRICE_NAME_COL, cboPriceNameCol.Text
    SaveSetting SET_PRICE_END_COL, cboPriceEndCol.Text
    SaveSetting SET_PRICE_UNIT_COL, cboPriceUnitCol.Text
    ThisWorkbook.Save: Unload Me
End Sub

Private Sub btnCancelPrice_Click()
    Unload Me
End Sub
```

---

## Step 16：ThisWorkbook のコード

**このステップでやること**：ブックの起動時・終了時の処理を記述します。

### 操作手順

1. VBEのプロジェクトエクスプローラーで `ThisWorkbook` をダブルクリック
2. コード編集エリアに以下のコードを貼り付けます

```vba
Option Explicit

'==================================================
' ThisWorkbook コード
' ブックの起動時・終了時の処理
'==================================================

'--------------------------------------------------
' ブックを開いたとき
'--------------------------------------------------

Private Sub Workbook_Open()
    ' 画面更新を一時停止
    Application.ScreenUpdating = False
    
    ' 简繁对照シートを非表示にする（ユーザーが誤って編集しないように）
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = Me.Worksheets(SHEET_SIMP_TRAD)
    If Not ws Is Nothing Then
        ws.Visible = xlSheetHidden
    End If
    On Error GoTo 0
    
    ' Settingsシートを非表示にする
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

'--------------------------------------------------
' ブックを閉じる前
'--------------------------------------------------

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' 自動保存
    On Error Resume Next
    Me.Save
    On Error GoTo 0
End Sub
```

> **重要**：`Workbook_Open` イベントは、マクロが有効な状態でブックを開いたときに自動実行されます。このコードにより、ブックを開くだけでメイン画面が表示されます。

---

## Step 17：動作テスト

**このステップでやること**：作成したツールが正しく動作するかテストします。

### 17-1. コンパイルチェック

1. VBEでメニュー「デバッグ」→「VBAプロジェクトのコンパイル」をクリック
2. エラーが表示されなければコンパイル成功です
3. エラーが表示された場合は、該当行を確認して修正します

### 17-2. 基本動作テスト

以下の手順でテストを実施します。

| テスト番号 | テスト項目 | 確認内容 |
|-----------|-----------|---------|
| 1 | ブックを開く | メイン画面（frmMain）が自動表示されること |
| 2 | 拠点選択 | cboLocation で北京/大連/河南が選択できること |
| 3 | フォルダ参照 | 「参照...」ボタンでフォルダ選択ダイアログが開くこと |
| 4 | ファイル一覧 | 選択したフォルダ内のExcelファイルが lstFiles に表示されること |
| 5 | 全选チェック | chkSelectAll で全ファイルが選択/解除されること |
| 6 | 選択件数 | ファイル選択に応じて lblCount が更新されること |
| 7 | プロジェクト設定 | btnProject で frmProject が開き、入力・保存できること |
| 8 | 人員設定 | btnPerson で frmPerson が開き、追加/更新/削除ができること |
| 9 | 単価表設定 | btnPrice で frmPrice が開き、設定・保存できること |
| 10 | 勤怠集計実行 | ファイルを選択して btnExecute で集計が実行されること |
| 11 | 3ファイル一括生成 | 集計後に btnGenerate で請求書が生成されること |
| 12 | 出力フォルダを開く | btnOpenFolder でエクスプローラーが開くこと |
| 13 | 設定の永続化 | ブックを閉じて再度開いたとき、前回の設定が復元されること |

### 17-3. テスト用データの準備

テストには以下のような簡易データを用意することをお勧めします。

1. テスト用の勤怠表ファイル（Excel）を1〜2件作成
   - 氏名行と日別データ（出勤時刻・退勤時刻）を含む
   - 例：`北京_202604_テスト勤怠.xlsx`
2. テスト用の単価表ファイルを作成
   - 氏名・退社日・契約単価の列を含む
3. テンプレートファイル3点を用意（空のExcelファイルでも可）

### 17-4. エラー発生時の確認ポイント

| 症状 | 確認ポイント |
|------|-------------|
| メイン画面が表示されない | マクロが有効か確認。セキュリティ設定で「すべてのマクロを有効にする」に設定 |
| 「コンパイルエラー」が表示される | 変数名や関数名のスペルミス。Option Explicit で宣言漏れがないか確認 |
| ファイル一覧が空 | フォルダパスが正しいか。フォルダ内に .xlsx / .xls ファイルがあるか確認 |
| 集計結果が0件 | 勤怠表のフォーマットが想定と異なる可能性。ParseBeijingFormat のロジックを確認 |
| 請求書が生成されない | テンプレートパスが正しく設定されているか。出力先フォルダの書き込み権限があるか確認 |
| 「実行時エラー 9」 | シート名が存在しない。Settings / 简繁对照 シートが正しく作成されているか確認 |
| 「実行時エラー 53」 | ファイルが見つからない。テンプレートや単価表のパスを再確認 |

---

## 補足：トラブルシューティング FAQ

### Q1. マクロが実行できません。「セキュリティの警告」が表示されます

**A.** Excelのマクロセキュリティ設定を変更します。
1. 「ファイル」→「オプション」→「セキュリティ センター」→「セキュリティ センターの設定」
2. 「マクロの設定」で「すべてのマクロを有効にする」を選択
3. 「OK」をクリック
4. ブックを一度閉じて再度開きます

> **注意**：この設定はセキュリティリスクを伴います。信頼できるブックのみ開くようにしてください。または、ブックを「信頼できる場所」に保存する方法もあります。

---

### Q2. 「コンパイルエラー：変数が定義されていません」と表示されます

**A.** 各モジュールの先頭に `Option Explicit` を記述しているため、すべての変数を宣言する必要があります。エラーメッセージで指摘された変数名を確認し、`Dim` ステートメントで宣言を追加してください。

---

### Q3. ファイル一覧に何も表示されません

**A.** 以下の点を確認してください。
- フォルダパスが正しいか（末尾に `\` は不要）
- フォルダ内に `.xlsx` または `.xls` ファイルが存在するか
- 担当者フィルター（cboPerson）が「(すべて)」になっているか
- 期間フィルター（txtPeriod）が空欄か、正しい形式（例：`2026/4`）か

---

### Q4. 勤怠集計を実行しても結果が0件です

**A.** 勤怠表のフォーマットが想定と異なる可能性があります。`modMain` の `ParseBeijingFormat` 関数が想定するフォーマット：
- 1列目に氏名（漢字2〜4文字）
- 氏名行の次の行から日別データ（1列目：日付、2列目：出勤時刻、3列目：退勤時刻）

実際の勤怠表フォーマットに合わせて `ParseBeijingFormat` のロジックを調整してください。

---

### Q5. 簡体字の氏名が正しくマッチングされません

**A.** `简繁对照` シートに変換マッピングが不足している可能性があります。
1. `简繁对照` シートを表示（右クリック → 「再表示」）
2. 不足している簡体字と対応する繁体字を追加
3. 再度シートを非表示にします

---

### Q6. 出力された請求書の金額が計算されていません

**A.** テンプレート側の計算式が正しく設定されているか確認してください。本ツールは生データ（氏名、工数、単価など）のみを埋め込み、金額計算はテンプレートのExcel計算式（`=E16*H16` など）に任せています。

---

### Q7. 一人が複数の作業を担当している場合、どう設定すればよいですか

**A.** 人員設定画面（frmPerson）で、同じ氏名に対して複数行を登録してください。
例：
- 行1：張三 / 2026/4/1〜4/15 / 開発 / W001 / Senior 1
- 行2：張三 / 2026/4/16〜4/30 / テスト / W002 / Senior 1

---

### Q8. 出力ファイル名をカスタマイズしたいです

**A.** `modMain` の `GenerateHSCN`、`GenerateCompanyReport`、`GenerateOrderReport` 関数内の出力ファイル名を変更してください。

```vba
' 変更前
outputFile = outputPath & "\HSCN验收书_" & Format(Now, "yyyymmdd") & ".xlsx"

' 変更後（例：プロジェクト名を含める）
outputFile = outputPath & "\HSCN验收书_" & GetSetting(SET_PROJECT_NAME) & "_" & Format(Now, "yyyymmdd") & ".xlsx"
```

---

### Q9. 昼休憩時間を1時間以外に変更したいです

**A.** `modConstants` の以下の定数を変更してください。

```vba
' 変更前
Public Const LUNCH_BREAK_HOURS As Double = 1#

' 変更後（例：45分 = 0.75時間）
Public Const LUNCH_BREAK_HOURS As Double = 0.75
```

---

### Q10. 勤務時間の丸め単位を0.5時間から変更したいです

**A.** `modConstants` の以下の定数を変更してください。

```vba
' 変更前
Public Const ROUND_UNIT_HOURS As Double = 0.5

' 変更後（例：15分単位 = 0.25時間）
Public Const ROUND_UNIT_HOURS As Double = 0.25
```

---

### Q11. 1人月の基準時間（168時間）を変更したいです

**A.** `modConstants` の以下の定数を変更してください。

```vba
' 変更前
Public Const HOURS_PER_MAN_MONTH As Double = 168#

' 変更後（例：160時間）
Public Const HOURS_PER_MAN_MONTH As Double = 160#
```

---

### Q12. 新しい拠点を追加したいです

**A.** 以下の2箇所を修正します。

1. `modConstants` に拠点定数を追加：
```vba
Public Const LOCATION_SHANGHAI As String = "上海"
```

2. `frmMain` の `UserForm_Initialize` でコンボボックスに追加：
```vba
cboLocation.AddItem LOCATION_SHANGHAI
```

3. `modMain` の `ProcessAttendanceFile` に新しい拠点の分岐を追加（必要に応じて専用の解析関数を作成）

---

### Q13. エラー「実行時エラー 1004：アプリケーション定義またはオブジェクト定義のエラーです」が発生します

**A.** このエラーは主に以下の原因で発生します。
- 存在しないシート名を参照している
- 保護されたシートに書き込もうとしている
- ファイルが読み取り専用で開かれている

Settingsシートのテンプレートパスが正しいか、テンプレートファイルが破損していないか確認してください。

---

### Q14. 処理が遅いです。高速化する方法はありますか

**A.** 以下の対策が有効です。
- 集計前に不要なアプリケーションを終了する
- `Application.ScreenUpdating = False` が各処理の冒頭で設定されているか確認
- 大量ファイルを処理する場合は、`Application.Calculation = xlCalculationManual` を追加（処理後に `xlCalculationAutomatic` に戻す）
- 単価表の読み込みを1回だけにしているか確認（`LoadPriceTable` は集計開始時に1回のみ呼ばれます）

---

### Q15. Mac版Excelで動作しますか

**A.** 本ツールはWindows版Excelを前提としています。Mac版では以下の制約があります。
- `FileSystemObject`（`Scripting.FileSystemObject`）が使用できない場合があります
- `Shell` 関数の動作が異なります
- フォルダ選択ダイアログ（`FileDialog`）の挙動が異なる場合があります

Macで使用する場合は、これらの部分をMac対応のコードに書き換える必要があります。

---

## 付録：全コード一覧（クイックリファレンス）

### モジュール構成

| モジュール名 | 種類 | 内容 |
|-------------|------|------|
| modConstants | 標準モジュール | 定数定義 |
| modUtils | 標準モジュール | ユーティリティ関数（簡繁体変換、時間計算、設定読み書き、ファイル操作） |
| modMain | 標準モジュール | メイン処理（勤怠集計、請求書生成） |
| frmMain | ユーザーフォーム | メイン画面 |
| frmProject | ユーザーフォーム | プロジェクト設定画面 |
| frmPerson | ユーザーフォーム | 人員設定画面 |
| frmPrice | ユーザーフォーム | 単価表設定画面 |
| ThisWorkbook | ブック | 起動時・終了時処理 |

### シート構成

| シート名 | 内容 | 表示 |
|---------|------|------|
| Settings | 設定値の保存（キー・バリュー形式） | 非表示（自動） |
| 简繁对照 | 簡体字→繁体字マッピング表 | 非表示（自動） |

---


### デバッグエラー　２０２６／０５／２８
集計エラーについて：
Private Sub ProcessAttendanceFile
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)　　ー＞filePath取得エラー。
    解決：GetSetting("LastLocation")でフォルダーパス結合

生成エラーについて： 
modMain:
GenerateAllReports 


> **本手順書について**
>
> 本手順書は、プログラミング未経験者でも画面キャプチャなしでVBAツールを再現できることを目的として作成されています。
> 各ステップのコードはコピー＆ペーストで使用できます。
> 実際の勤怠表フォーマットやテンプレートのレイアウトに合わせて、適宜コードを調整してください。
>
> 作成日：2026年5月27日
> バージョン：v2.0
