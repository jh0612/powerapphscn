Attribute VB_Name = "modConstants"
Option Explicit

'==================================================
' 定数定義モジュール
' 勤怠集計ツールで使用する定数を一元管理します
'==================================================

' --- シート名定数 ---
Public Const SHEET_SETTINGS As String = "Settings"
Public Const SHEET_SIMP_TRAD As String = "簡繁対照"

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
Public Const SERVICE_MODE_TM As String = "T&M（工?与材料）"
Public Const SERVICE_MODE_LUMPSUM As String = "FIXED PRICE（固定?用）"

'--- 出力ファイル名 ---
Public Const OUTFILE_MEISAI As String = "明細書"
Public Const OUTFILE_HSCN_ACCEPTANCE_REPORT As String = "HSCN験収書"
Public Const OUTFILE_HSCN_PO As String = "発注書"

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
Public Const COL_WEEKDAY_OT As Long = 10      ' 平日残業時間（合計）
Public Const COL_WEEKEND_OT As Long = 11      ' 休日残業時間（合計）

' --- 案件名設定キー（明細書・発注書のフォールバック用） ---
Public Const SET_CASE_NAME As String = "CaseName"  ' 案件名
