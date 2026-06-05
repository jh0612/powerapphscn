Attribute VB_Name = "modMain"
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
        'If Len(periodFilter) > 0 Then
            'If Not IsFileInPeriod(filePath, periodFilter) Then
                'GoTo NextFile
            'End If
        'End If
        
        ' 勤怠ファイルを読み取り
        Call ProcessAttendanceFile(filePath, location, priceDict)
        
NextFile:
    Next i
    
    ExecuteAggregation = True
    Exit Function
    
ErrHandler:
    LogToListBox "ExecuteAggregation Error: " & Err.Description
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
    headerRow = CLng(val(GetSetting(SET_PRICE_HEADER_ROW)))
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
    lastRow = ws.Cells(ws.Rows.count, nameCol).End(xlUp).Row
    
    Dim r As Long
    For r = headerRow + 1 To lastRow
        Dim rawName As String
        rawName = Trim(CStr(ws.Cells(r, nameCol).value))
        If Len(rawName) > 0 Then
            Dim normName As String
            normName = NormalizeName(rawName)
            
            Dim priceInfo(0 To 2) As Variant
            priceInfo(0) = rawName  ' 元の氏名
            priceInfo(1) = ws.Cells(r, unitCol).value  ' 契約単価
            priceInfo(2) = ws.Cells(r, endCol).value   ' 退社日
            
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
' 1つの勤怠ファイルを処理する
Private Sub ProcessAttendanceFile(ByVal filePath As String, _
                                   ByVal location As String, _
                                   ByVal priceDict As Object)
    On Error GoTo ErrorHandler
    
    ' ファイルの存在確認
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FileExists(filePath) Then
        Debug.Print "File not found (ProcessAttendanceFile): " & filePath
        Exit Sub
    End If
    
    Dim wb As Workbook
    Application.ScreenUpdating = False
    
    ' Unicodeファイル名に対応するため、FSO経由でファイルを開く
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False)
    
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)
    
    ' 拠点別にフォーマットが異なるため、拠点に応じた読み取りを行う
    Select Case location
        Case LOCATION_BEIJING
            Call ParseBeijingFormat(filePath, ws, priceDict)
        Case LOCATION_DALIAN
            Call ParseDalianFormat(filePath, ws, priceDict)
        Case LOCATION_HENAN
            Call ParseHenanFormat(filePath, ws, priceDict)
        Case Else
            ' デフォルト：北京形式として処理
            Call ParseBeijingFormat(filePath, ws, priceDict)
    End Select
    
    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Debug.Print "ProcessAttendanceFile Error: " & Err.Description & " - File: " & filePath
    If Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Application.ScreenUpdating = True
End Sub

'--------------------------------------------------
' 拠点別フォーマット解析
'--------------------------------------------------

''' 北京拠点の勤怠フォーマットを解析
''' 想定レイアウト：
'''   Row 1：ヘッダー（会社名など）
'''   Row 2：日別データ
Private Sub ParseBeijingFormat(ByVal filePath As String, ByVal ws As Worksheet, ByVal priceDict As Object)
    ' 簡易実装：全セルをスキャンして氏名行と日付行を検出
    
    Dim companyName As String
    Dim fileName As String
    companyName = Trim(CStr(ws.Cells(1, 1).value))
    fileName = filePath
    
    
    
    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    Dim r As Long
    For r = 2 To lastRow  '20260528 add
        Dim firstCell As String
        Dim regEx As Object
        Dim matchList As Object
        
        ' 正規表現声明
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Pattern = "/(?<=?力-)(.*?)(?=-)/g"
        regEx.Global = False
        regEx.Multiline = False
        
        
        ' 正規表現マッチ
        ' Set matchList = regEx.Execute(fileName)
    
        If regEx.Test(fileName) Then
            firstCell = regEx.Execute(fileName)(0).SubMatches(0)  ' 直接拿到人名
            MsgBox "提取到的人名：" & firstCell
        Else
            MsgBox "未匹配到人名"
        End If
        
        
        
        
        
        
        
        
        
        
        
        
        'firstCell = Trim(CStr(ws.Cells(r, 3).value))
        
        ' 氏名行の検出（3列目の氏名）ファイル名正規表現
       ' If IsNameCell(firstCell) Then
            Dim rawPersonName As String
            rawPersonName = firstCell
            Dim normPersonName As String
            normPersonName = NormalizeName(rawPersonName)
            
            ' 日別データの読み取り
            Dim dailyDict As Object
            Set dailyDict = CreateObject("Scripting.Dictionary")
            
            Dim totalHours As Double
            totalHours = 0
            
            ' 2行目以降を日別データとして読み取り
            Dim dr As Long
            dr = 2
            Do While dr <= lastRow
                Dim dayCell As String
                dayCell = Trim(CStr(ws.Cells(dr, 4).value)) '日期
                
                ' 次の氏名行に当たったら終了
                If IsNameCell(dayCell) Then Exit Do
                
                ' 日付データがあれば処理
                If IsNumeric(ToDate(dayCell)) And Len(ToDate(dayCell)) > 0 Then
                    Dim dayNum As Long
                    dayNum = CLng(ToDate(dayCell))
                    
                    Dim timeInVal As Variant, timeOutVal As Variant
                    timeInVal = ws.Cells(dr, 5).value  ' 出勤時刻（5列目想定）
                    timeOutVal = ws.Cells(dr, 7).value ' 退勤時刻（7列目想定）
                    
                    If IsDate(timeInVal) And IsDate(timeOutVal) Then
                        Dim workHrs As Double, overtime As Double
                        workHrs = CalcWorkHours(CDate(timeInVal), CDate(timeOutVal))
                        overtime = CalcOvertimeHours(workHrs)
                        
                        Dim dayInfo(0 To 4) As Variant
                        dayInfo(0) = dayNum
                        dayInfo(1) = Format(timeInVal, "hh:mm")
                        dayInfo(2) = Format(timeOutVal, "hh:mm")
                        dayInfo(3) = workHrs
                        'add
                        dayInfo(4) = overtime
                        
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
            '会社名マッピング
            Dim arrRow, arrCol, i&
           arrRow = Split(GetSetting("personData"), vbCrLf)
           For i = LBound(arrRow) To UBound(arrRow)
            If Trim(arrRow(i)) <> "" Then
              arrCol = Split(arrRow(i), vbTab)
              If UBound(arrCol) >= 5 And Trim(arrCol(0)) = Trim(rawPersonName) Then
             
               summary(COL_WORK_CONTENT) = arrCol(2) '作業内容
               summary(COL_COMPANY) = arrCol(3) '会社名
                summary(COL_WORK_NO) = arrCol(4) '作業番号
                 summary(COL_LEVEL) = arrCol(5)  'レベル
               Exit For
             End If
               End If
            Next i
            
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
       ' End If
    Next r
End Sub

''' 大連拠点の勤怠フォーマットを解析（北京と同じ基本構造を想定）
Private Sub ParseDalianFormat(ByVal filePath As String, ByVal ws As Worksheet, ByVal priceDict As Object)
    ' 大連形式：北京と同様の構造を想定
    Call ParseBeijingFormat(ws, priceDict)
End Sub

''' 河南拠点の勤怠フォーマットを解析（北京と同じ基本構造を想定）
Private Sub ParseHenanFormat(ByVal filePath As String, ByVal ws As Worksheet, ByVal priceDict As Object)
    ' 河南形式：北京と同様の構造を想定
    Call ParseBeijingFormat(ws, priceDict)
End Sub

''' セルの値が氏名かどうかを判定する
''' 簡易判定：漢字2?4文字で数字を含まない
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
Public Function IsFileInPeriod(ByVal filePath As String, ByVal period As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim fileName As String
    fileName = CStr(fso.GetFileName(filePath))
    
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
    
    ' 1. HSCN験収書生成
    If Not GenerateHSCN(outputPath) Then
        LogToListBox "HSCN験収書生成に失敗しました"
    End If
    
    ' 2. 明細書生成
    If Not GenerateCompanyReport(outputPath) Then
        LogToListBox "明細書生成に失敗しました"
    End If
    
    ' 3. 発注書生成
    If Not GenerateOrderReport(outputPath) Then
        LogToListBox "発注書生成に失敗しました"
    End If
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    GenerateAllReports = True
    Exit Function
    
ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    LogToListBox "GenerateAllReports Error: " & Err.Description
    GenerateAllReports = False
End Function

''' HSCN?收?を生成する
Private Function GenerateHSCN(ByVal outputPath As String) As Boolean
    Dim templatePath As String
    templatePath = GetSetting(SET_HSCN_PATH)
    
    If Len(templatePath) = 0 Or Len(Dir(templatePath)) = 0 Then
        MsgBox OUTFILE_HSCN_ACCEPTANCE_REPORT & "のテンプレートが見つかりません。", vbExclamation
        GenerateHSCN = False
        Exit Function
    End If
    
    ' 出力ファイル名
    Dim outputFile As String
    outputFile = outputPath & "\" & OUTFILE_HSCN_ACCEPTANCE_REPORT & "_" & Format(Now, "yyyymmddhhmmss ") & ".xlsx"
    
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
            templateWS.Copy After:=wb.Worksheets(wb.Worksheets.count)
            Dim newWS As Worksheet
            Set newWS = ActiveSheet  'wb.Worksheets(wb.Worksheets.count)
            
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
            newWS.Cells(3, 3).value = CStr(summary(COL_COMPANY))
            
            ' Row 6：氏名、等級
            newWS.Cells(4, 3).value = personName
            newWS.Cells(4, 5).value = CStr(summary(COL_LEVEL))
            
            ' Row 8：年月、総工時、人月
            newWS.Cells(5, 4).value = ConvertToDateStr(frmMain.txtPeriod.Text)
            'newWS.Cells(8, 4).value = CDbl(summary(COL_TOTAL_HOURS))
            'newWS.Cells(8, 6).value = CDbl(summary(COL_MAN_MONTH))
            
            '帳票対象期間文字列より年月取得
            Dim baseDate As Date
              baseDate = ConvertToDateStr(frmMain.txtPeriod.Text)
            Dim baseY As Integer
            baseY = Year(baseDate)
            
            ' Row 11以降：日別データ
            Dim dailyDict As Object
            Set dailyDict = summary(COL_DAILY_DATA)
            
            If Not dailyDict Is Nothing And dailyDict.count > 0 Then
                Dim rowIdx As Long
                rowIdx = 8
                
                Dim dk As Variant
                For Each dk In dailyDict.Keys
                    Dim dayInfo As Variant
                    dayInfo = dailyDict(dk)
                    
                    'newWS.Cells(rowIdx, 1).value = month(dayInfo(0))  ' 月
                    'newWS.Cells(rowIdx, 2).value = Day(dayInfo(0))   ' 日
                    newWS.Cells(rowIdx, 3).value = GetWorkContentByDate(personName, newWS.Cells(8, 1).value, day(newWS.Cells(rowIdx, 2).value), baseY)      'CStr(summary(COL_WORK_CONTENT)) 'GetSetting(SET_SERVICE_CONTENT)           ' 作業内容（後で手動入力）
                    newWS.Cells(rowIdx, 5).value = dayInfo(1)   ' 出勤
                    newWS.Cells(rowIdx, 6).value = dayInfo(2)   ' 退勤
                    newWS.Cells(rowIdx, 7).value = dayInfo(4)            ' 残業班（テンプレの計算式想定）
                    newWS.Cells(rowIdx, 8).value = IIf(dayInfo(3) < 8, 8 - dayInfo(3), 0)          ' 欠勤
                    newWS.Cells(rowIdx, 9).value = IIf(dayInfo(3) >= 8, 8, dayInfo(3))  ' 工数
                    
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
        ' Row 5プロジェクト設定から埋め込み
        accWS.Cells(5, 2).value = GetSetting(SET_PROJECT_NAME)
        accWS.Cells(6, 2).value = GetSetting(SET_SUPPLIER)
        accWS.Cells(7, 2).value = GetSetting(SET_SERVICE_MODE)
        accWS.Cells(8, 2).value = GetSetting(SET_SERVICE_CONTENT)
        'accWS.Cells(9, 2).value = GetSetting(SET_DELIVERABLE)
        accWS.Cells(10, 2).value = GetSetting(SET_MILESTONE)
        accWS.Cells(11, 2).value = ConvertToDateStr(frmMain.txtPeriod.Text)   'GetSetting(SET_CONTRACT_NO)
        accWS.Cells(11, 4).value = ConvertToDateStr(frmMain.txtPeriodTo.Text)   'GetSetting(SET_PO)
        
        ' Row 15以降：人員別集計データ
        Dim accRow As Long
        accRow = 15
        Dim idx As Long
        idx = 1
        
        For Each key In m_summaryData.Keys
            summary = m_summaryData(key)
            
            'accWS.Cells(accRow, 1).value = idx
            accWS.Cells(accRow, 2).value = CStr(summary(COL_PERSON_NAME)) 'Name:
            accWS.Cells(accRow, 3).value = CStr(summary(COL_LEVEL)) 'level
            'accWS.Cells(accRow, 4).value = CStr(summary(COL_WORK_CONTENT))
           ' accWS.Cells(accRow, 6).value = CDbl(summary(COL_MAN_MONTH))
            accWS.Cells(accRow, 7).Formula = "=Timesheet_" & accWS.Cells(accRow, 2).value & "!I5"
            accWS.Cells(accRow, 9).value = CDbl(summary(COL_UNIT_PRICE))
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
        MsgBox OUTFILE_MEISAI & "のテンプレートが見つかりません。", vbExclamation
        GenerateCompanyReport = False
        Exit Function
    End If
    
    Dim outputFile As String
    outputFile = outputPath & "\" & OUTFILE_MEISAI & "_" & Format(Now, "yyyymmddhhmmss ") & ".xlsx"
    
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
        
        'ws.Cells(dataRow, 1).value = summary(COL_DAILY_DATA)  '日付
        ws.Cells(dataRow, 2).value = CStr(summary(COL_WORK_CONTENT))   '作業内容CStr(summary(COL_PERSON_NAME))
        ws.Cells(dataRow, 3).value = "公共シ" 'CStr(summary(COL_WORK_PERIOD))
        ws.Cells(dataRow, 4).value = CStr(summary(COL_WORK_CONTENT))  '案件内容
        ws.Cells(dataRow, 5).value = CDbl(summary(COL_TOTAL_HOURS))
        ws.Cells(dataRow, 6).value = CStr(summary(COL_PERSON_NAME))  '名前 CDbl(summary(COL_MAN_MONTH))
        ws.Cells(dataRow, 7).value = CDbl(summary(COL_UNIT_PRICE)) '単価　CStr(summary(COL_LEVEL))
        'ws.Cells(dataRow, 8).value = CDbl(summary(COL_UNIT_PRICE))
        
        dataRow = dataRow + 1
        idx = idx + 1
    Next key
    
    wb.Close SaveChanges:=True
    GenerateCompanyReport = True
End Function

''' 発注書を生成する
Private Function GenerateOrderReport(ByVal outputPath As String) As Boolean
    Dim templatePath As String
    templatePath = GetSetting(SET_ORDER_PATH)
    
    If Len(templatePath) = 0 Or Len(Dir(templatePath)) = 0 Then
        MsgBox OUTFILE_HSCN_PO & "のテンプレートが見つかりません。", vbExclamation
        GenerateOrderReport = False
        Exit Function
    End If
    
    Dim outputFile As String
    outputFile = outputPath & "\" & OUTFILE_HSCN_PO & "_" & Format(Now, "yyyymmddhhmmss ") & ".xls"
    
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
        
        ws.Cells(dataRow, 1).value = idx
        ws.Cells(dataRow, 2).value = CStr(summary(COL_PERSON_NAME))
        ws.Cells(dataRow, 3).value = CStr(summary(COL_WORK_PERIOD))
        ws.Cells(dataRow, 4).value = CStr(summary(COL_WORK_CONTENT))
        ws.Cells(dataRow, 5).value = CDbl(summary(COL_TOTAL_HOURS))
        ws.Cells(dataRow, 6).value = CDbl(summary(COL_MAN_MONTH))
        ws.Cells(dataRow, 7).value = CStr(summary(COL_LEVEL))
        ws.Cells(dataRow, 8).value = CDbl(summary(COL_UNIT_PRICE))
        
        dataRow = dataRow + 1
        idx = idx + 1
    Next key
    
    wb.Close SaveChanges:=True
    GenerateOrderReport = True
End Function

