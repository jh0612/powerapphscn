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
    ' File name format: e.g. name-company-personName-code-workno-startdate-enddate
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fileName As String
    fileName = fso.GetBaseName(filePath)

    Dim nameParts() As String
    nameParts = Split(fileName, "-")
    Dim rawPersonName As String
    If UBound(nameParts) >= 2 Then
        rawPersonName = Trim(nameParts(2))
    Else
        LogToListBox "Cannot extract person name from: " & fileName
        Exit Sub
    End If
    Dim normPersonName As String
    normPersonName = NormalizeName(rawPersonName)

    Dim baseDate As Date
    On Error Resume Next
    baseDate = CDate(ConvertToDateStr(frmMain.txtPeriod.Text))
    On Error GoTo 0
    Dim baseY As Integer
    If Year(baseDate) > 2000 Then
        baseY = Year(baseDate)
    Else
        baseY = Year(Now)
    End If

    Dim dailyDict As Object
    Set dailyDict = CreateObject("Scripting.Dictionary")
    Dim totalHours As Double
    Dim totalWeekdayOT As Double
    Dim totalWeekendOT As Double
    totalHours = 0#
    totalWeekdayOT = 0#
    totalWeekendOT = 0#

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row

    Dim dr As Long
    For dr = 2 To lastRow
        Dim dayCell As String
        dayCell = Trim(CStr(ws.Cells(dr, 4).Value))

        Dim dateStr As String
        dateStr = ToDate(dayCell)

        If IsValid8Date(dateStr) Then
            Dim timeInVal As Variant, timeOutVal As Variant
            timeInVal = ws.Cells(dr, 5).Value
            timeOutVal = ws.Cells(dr, 7).Value

            If IsDate(timeInVal) And IsDate(timeOutVal) Then
                Dim workHrs As Double
                workHrs = CalcWorkHours(CDate(timeInVal), CDate(timeOutVal))

                Dim dateVal As Date
                dateVal = DateSerial(CLng(Left(dateStr, 4)), CLng(Mid(dateStr, 5, 2)), CLng(Right(dateStr, 2)))
                Dim isWeekend As Boolean
                isWeekend = (Weekday(dateVal) = vbSunday Or Weekday(dateVal) = vbSaturday)

                Dim wdOT As Double, weOT As Double, otHrs As Double, stdHrs As Double
                If isWeekend Then
                    stdHrs = 0#
                    weOT = workHrs
                    wdOT = 0#
                    otHrs = workHrs
                Else
                    stdHrs = IIf(workHrs >= STANDARD_WORK_HOURS, STANDARD_WORK_HOURS, workHrs)
                    otHrs = CalcOvertimeHours(workHrs)
                    wdOT = otHrs
                    weOT = 0#
                End If

                totalHours = totalHours + stdHrs
                totalWeekdayOT = totalWeekdayOT + wdOT
                totalWeekendOT = totalWeekendOT + weOT

                Dim dayInfo(0 To 6) As Variant
                dayInfo(0) = dateStr
                dayInfo(1) = Format(timeInVal, "hh:mm")
                dayInfo(2) = Format(timeOutVal, "hh:mm")
                dayInfo(3) = workHrs
                dayInfo(4) = otHrs
                dayInfo(5) = wdOT
                dayInfo(6) = weOT

                dailyDict.Add CStr(dr), dayInfo
            End If
        End If
    Next dr

    Dim manMonth As Double
    manMonth = CalcManMonth(totalHours)

    Dim unitPrice As Variant
    unitPrice = 0
    If priceDict.Exists(normPersonName) Then
        Dim pInfo As Variant
        pInfo = priceDict(normPersonName)
        unitPrice = pInfo(1)
    End If

    Dim summary(0 To 11) As Variant
    summary(COL_PERSON_NAME) = rawPersonName
    summary(COL_LEVEL) = ""
    summary(COL_WORK_PERIOD) = ""
    summary(COL_WORK_CONTENT) = ""
    summary(COL_WORK_NO) = ""
    summary(COL_TOTAL_HOURS) = totalHours
    summary(COL_MAN_MONTH) = manMonth
    summary(COL_UNIT_PRICE) = unitPrice
    summary(COL_COMPANY) = Trim(CStr(ws.Cells(1, 1).Value))
    Set summary(COL_DAILY_DATA) = dailyDict
    summary(COL_WEEKDAY_OT) = totalWeekdayOT
    summary(COL_WEEKEND_OT) = totalWeekendOT

    Dim arrRow As Variant, arrCol As Variant, i As Long
    arrRow = Split(GetSetting("personData"), vbCrLf)
    For i = LBound(arrRow) To UBound(arrRow)
        If Trim(arrRow(i)) <> "" Then
            arrCol = Split(arrRow(i), vbTab)
            If UBound(arrCol) >= 5 And Trim(arrCol(0)) = Trim(rawPersonName) Then
                summary(COL_WORK_PERIOD) = arrCol(1)
                summary(COL_WORK_CONTENT) = arrCol(2)
                summary(COL_COMPANY) = arrCol(3)
                summary(COL_WORK_NO) = arrCol(4)
                summary(COL_LEVEL) = arrCol(5)
                Exit For
            End If
        End If
    Next i

    If m_summaryData.Exists(normPersonName) Then
        Dim existing As Variant
        existing = m_summaryData(normPersonName)
        existing(COL_TOTAL_HOURS) = CDbl(existing(COL_TOTAL_HOURS)) + totalHours
        existing(COL_MAN_MONTH) = CalcManMonth(CDbl(existing(COL_TOTAL_HOURS)))
        existing(COL_WEEKDAY_OT) = CDbl(existing(COL_WEEKDAY_OT)) + totalWeekdayOT
        existing(COL_WEEKEND_OT) = CDbl(existing(COL_WEEKEND_OT)) + totalWeekendOT
        m_summaryData(normPersonName) = existing
    Else
        m_summaryData.Add normPersonName, summary
    End If
End Sub

Private Sub ParseDalianFormat(ByVal filePath As String, ByVal ws As Worksheet, ByVal priceDict As Object)
    Call ParseBeijingFormat(filePath, ws, priceDict)
End Sub

Private Sub ParseHenanFormat(ByVal filePath As String, ByVal ws As Worksheet, ByVal priceDict As Object)
    Call ParseBeijingFormat(filePath, ws, priceDict)
End Sub

Private Function IsNameCell(ByVal cellValue As String) As Boolean
    If Len(cellValue) < 2 Or Len(cellValue) > 4 Then
        IsNameCell = False
        Exit Function
    End If
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

Public Function IsFileInPeriod(ByVal filePath As String, ByVal period As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim fileName As String
    fileName = CStr(fso.GetFileName(filePath))
    Dim periodVariants(0 To 3) As String
    periodVariants(0) = Replace(period, "/", "")
    periodVariants(1) = Replace(period, "/", "_")
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

Private Function GetCaseNameByWorkNo(ByVal workNo As String) As String
    If Len(Trim(workNo)) = 0 Then
        GetCaseNameByWorkNo = GetSetting(SET_CASE_NAME)
        Exit Function
    End If
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("name")
    On Error GoTo 0
    If ws Is Nothing Then
        GetCaseNameByWorkNo = GetSetting(SET_CASE_NAME)
        Exit Function
    End If
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row
    Dim i As Long
    For i = 2 To lastRow
        If Trim(CStr(ws.Cells(i, 12).Value)) = Trim(workNo) Then
            GetCaseNameByWorkNo = Trim(CStr(ws.Cells(i, 13).Value))
            Exit Function
        End If
    Next i
    GetCaseNameByWorkNo = GetSetting(SET_CASE_NAME)
End Function

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

    Dim outputFile As String
    outputFile = outputPath & "\\" & OUTFILE_HSCN_ACCEPTANCE_REPORT & "_" & Format(Now, "yyyymmddhhmmss") & ".xlsx"
    FileCopy templatePath, outputFile

    Dim wb As Workbook
    Set wb = Workbooks.Open(outputFile)

    Dim templateWS As Worksheet
    On Error Resume Next
    Set templateWS = wb.Worksheets(TEMPLATE_TIMESHEET)
    On Error GoTo 0

    If Not templateWS Is Nothing Then
        Dim key As Variant
        For Each key In m_summaryData.Keys
            Dim summary As Variant
            summary = m_summaryData(key)

            templateWS.Copy After:=wb.Worksheets(wb.Worksheets.Count)
            Dim newWS As Worksheet
            Set newWS = ActiveSheet

            Dim personName As String
            personName = CStr(summary(COL_PERSON_NAME))

            Dim sheetName As String
            sheetName = "Timesheet_" & personName
            If Len(sheetName) > 31 Then sheetName = Left(sheetName, 31)
            On Error Resume Next
            newWS.Name = sheetName
            On Error GoTo 0

            newWS.Cells(3, 3).Value = CStr(summary(COL_COMPANY))
            newWS.Cells(4, 3).Value = personName
            newWS.Cells(4, 5).Value = CStr(summary(COL_LEVEL))
            newWS.Cells(5, 4).Value = ConvertToDateStr(frmMain.txtPeriod.Text)

            Dim dailyDict As Object
            Set dailyDict = summary(COL_DAILY_DATA)

            If Not dailyDict Is Nothing And dailyDict.Count > 0 Then
                Dim rowIdx As Long
                rowIdx = 8
                Dim dk As Variant
                For Each dk In dailyDict.Keys
                    Dim dayInfo As Variant
                    dayInfo = dailyDict(dk)

                    Dim dateStr As String
                    dateStr = CStr(dayInfo(0))
                    Dim dateM As Long, dateD As Long
                    dateM = CLng(Mid(dateStr, 5, 2))
                    dateD = CLng(Right(dateStr, 2))
                    newWS.Cells(rowIdx, 1).Value = dateM
                    newWS.Cells(rowIdx, 2).Value = dateD

                    Dim dateY As Long
                    dateY = CLng(Left(dateStr, 4))
                    newWS.Cells(rowIdx, 3).Value = GetWorkContentByDate(personName, dateM, dateD, dateY)

                    newWS.Cells(rowIdx, 5).Value = dayInfo(1)
                    newWS.Cells(rowIdx, 6).Value = dayInfo(2)
                    newWS.Cells(rowIdx, 7).Value = dayInfo(4)

                    Dim stdH As Double
                    Dim isWE As Boolean
                    isWE = CDbl(dayInfo(6)) > 0 And CDbl(dayInfo(5)) = 0 And CDbl(dayInfo(3)) > 0
                    If isWE Then
                        stdH = 0#
                    Else
                        stdH = IIf(dayInfo(3) >= STANDARD_WORK_HOURS, STANDARD_WORK_HOURS, dayInfo(3))
                    End If
                    newWS.Cells(rowIdx, 8).Value = IIf(stdH < STANDARD_WORK_HOURS And Not isWE, STANDARD_WORK_HOURS - stdH, 0)
                    newWS.Cells(rowIdx, 9).Value = stdH

                    rowIdx = rowIdx + 1
                Next dk
            End If
        Next key

        templateWS.Visible = xlSheetHidden
    End If

    Dim accWS As Worksheet
    On Error Resume Next
    Set accWS = wb.Worksheets(TEMPLATE_ACCEPTANCE)
    On Error GoTo 0

    If Not accWS Is Nothing Then
        accWS.Cells(5, 2).Value = GetSetting(SET_PROJECT_NAME)
        accWS.Cells(6, 2).Value = GetSetting(SET_SUPPLIER)
        accWS.Cells(7, 2).Value = GetSetting(SET_SERVICE_MODE)
        accWS.Cells(8, 2).Value = GetSetting(SET_SERVICE_CONTENT)
        accWS.Cells(10, 2).Value = GetSetting(SET_MILESTONE)
        accWS.Cells(11, 2).Value = ConvertToDateStr(frmMain.txtPeriod.Text)
        accWS.Cells(11, 4).Value = ConvertToDateStr(frmMain.txtPeriodTo.Text)

        Dim accRow As Long
        accRow = 15

        For Each key In m_summaryData.Keys
            summary = m_summaryData(key)
            personName = CStr(summary(COL_PERSON_NAME))

            Dim pStart As String, pEnd As String
            Dim wp As String
            wp = CStr(summary(COL_WORK_PERIOD))
            If Len(wp) > 0 Then
                Dim wpParts() As String
                wpParts = Split(wp, "-")
                If UBound(wpParts) >= 1 Then
                    pStart = Trim(wpParts(0))
                    pEnd = Trim(wpParts(1))
                End If
            End If
            If Len(pStart) = 0 Then pStart = ConvertToDateStr(frmMain.txtPeriod.Text)
            If Len(pEnd) = 0 Then pEnd = ConvertToDateStr(frmMain.txtPeriodTo.Text)

            accWS.Cells(accRow, 2).Value = personName
            accWS.Cells(accRow, 3).Value = CStr(summary(COL_LEVEL))
            accWS.Cells(accRow, 4).Value = pStart
            accWS.Cells(accRow, 6).Value = pEnd
            Dim tsName As String
            tsName = "Timesheet_" & personName
            If Len(tsName) > 31 Then tsName = Left(tsName, 31)
            accWS.Cells(accRow, 7).Formula = "=" & tsName & "!I15"
            accWS.Cells(accRow, 9).Value = CDbl(summary(COL_UNIT_PRICE))
            accWS.Cells(accRow, 10).Value = 1
            accRow = accRow + 1

            Dim wdOT As Double, weOT As Double
            wdOT = CDbl(summary(COL_WEEKDAY_OT))
            weOT = CDbl(summary(COL_WEEKEND_OT))
            If wdOT + weOT > 0 Then
                accWS.Cells(accRow, 2).Value = personName
                accWS.Cells(accRow, 3).Value = CStr(summary(COL_LEVEL))
                accWS.Cells(accRow, 4).Value = pStart
                accWS.Cells(accRow, 6).Value = pEnd
                accWS.Cells(accRow, 8).Value = "1式"
                accWS.Cells(accRow, 10).Value = 1
                accWS.Cells(accRow, 11).Value = wdOT * 15 + weOT * 20
                accRow = accRow + 1
            End If
        Next key
    End If

    wb.Close SaveChanges:=True
    GenerateHSCN = True
End Function

Private Function GenerateCompanyReport(ByVal outputPath As String) As Boolean
    Dim templatePath As String
    templatePath = GetSetting(SET_COMPANY_PATH)

    If Len(templatePath) = 0 Or Len(Dir(templatePath)) = 0 Then
        MsgBox OUTFILE_MEISAI & "のテンプレートが見つかりません。", vbExclamation
        GenerateCompanyReport = False
        Exit Function
    End If

    Dim outputFile As String
    outputFile = outputPath & "\\" & OUTFILE_MEISAI & "_" & Format(Now, "yyyymmddhhmmss") & ".xlsx"
    FileCopy templatePath, outputFile

    Dim wb As Workbook
    Set wb = Workbooks.Open(outputFile)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)

    Dim dataRow As Long
    dataRow = 2

    Dim key As Variant
    For Each key In m_summaryData.Keys
        Dim summary As Variant
        summary = m_summaryData(key)

        Dim workPeriod As String
        workPeriod = CStr(summary(COL_WORK_PERIOD))
        If Len(workPeriod) = 0 Then
            workPeriod = ConvertToDateStr(frmMain.txtPeriod.Text) & "~" & ConvertToDateStr(frmMain.txtPeriodTo.Text)
        End If
        ws.Cells(dataRow, 1).Value = workPeriod

        Dim workContent As String
        workContent = CStr(summary(COL_WORK_CONTENT))
        If Len(workContent) = 0 Then workContent = GetSetting(SET_SERVICE_CONTENT)
        ws.Cells(dataRow, 2).Value = workContent

        ws.Cells(dataRow, 3).Value = ""

        ws.Cells(dataRow, 4).Value = GetCaseNameByWorkNo(CStr(summary(COL_WORK_NO)))

        ws.Cells(dataRow, 5).Value = ""

        ws.Cells(dataRow, 6).Value = CStr(summary(COL_PERSON_NAME))

        ws.Cells(dataRow, 7).Value = CDbl(summary(COL_UNIT_PRICE))

        ws.Cells(dataRow, 8).Value = CDbl(summary(COL_TOTAL_HOURS))

        dataRow = dataRow + 1
    Next key

    wb.Close SaveChanges:=True
    GenerateCompanyReport = True
End Function

Private Function GenerateOrderReport(ByVal outputPath As String) As Boolean
    Dim templatePath As String
    templatePath = GetSetting(SET_ORDER_PATH)

    If Len(templatePath) = 0 Or Len(Dir(templatePath)) = 0 Then
        MsgBox OUTFILE_HSCN_PO & "のテンプレートが見つかりません。", vbExclamation
        GenerateOrderReport = False
        Exit Function
    End If

    Dim outputFile As String
    outputFile = outputPath & "\\" & OUTFILE_HSCN_PO & "_" & Format(Now, "yyyymmddhhmmss") & ".xlsx"
    FileCopy templatePath, outputFile

    Dim wb As Workbook
    Set wb = Workbooks.Open(outputFile)
    Dim ws As Worksheet
    Set ws = wb.Worksheets(1)

    Dim dataRow As Long
    dataRow = 13

    Dim key As Variant
    For Each key In m_summaryData.Keys
        Dim summary As Variant
        summary = m_summaryData(key)

        ws.Cells(dataRow, 2).Value = GetCaseNameByWorkNo(CStr(summary(COL_WORK_NO)))

        ws.Cells(dataRow, 6).Value = CDbl(summary(COL_TOTAL_HOURS))

        ws.Cells(dataRow, 7).Value = CDbl(summary(COL_UNIT_PRICE))

        ws.Cells(dataRow, 9).Value = "要員"

        ws.Cells(dataRow, 14).Value = CStr(summary(COL_PERSON_NAME))

        ws.Cells(dataRow, 15).Value = CStr(summary(COL_WORK_NO))

        dataRow = dataRow + 1
    Next key

    wb.Close SaveChanges:=True
    GenerateOrderReport = True
End Function
