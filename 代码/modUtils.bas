Attribute VB_Name = "modUtils"
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

''' 簡繁体変換辞書を簡繁照シートから読み込む
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
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 2 To lastRow  ' 1行目は見出し想定
        Dim simp As String, trad As String
        simp = Trim(CStr(ws.Cells(i, 1).value))
        trad = Trim(CStr(ws.Cells(i, 2).value))
        If Len(simp) > 0 And Len(trad) > 0 Then
            If Not m_dictSimplToTrad.Exists(simp) Then
                m_dictSimplToTrad.Add simp, trad
            End If
        End If
    Next i
    
    Set GetSimplTradDict = m_dictSimplToTrad
End Function

''' 簡体字を含む文字列を繁体字に正規化する
''' 簡繁照シートのマッピングに従って1文字ずつ変換します
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
''' 時刻を0.5時間単位で切り捨てる（上方丸め）
''' 例：9:16 → 9:30、18:29 → 18:30
Public Function RoundUpOnWork(ByVal dt As Date) As Double
    Dim hours As Double
    hours = Hour(dt) + Minute(dt) / 60#
    '切り上げ
    RoundUpOnWork = Application.WorksheetFunction.Ceiling(hours, ROUND_UNIT_HOURS)
End Function

''' 値を0.5時間単位で切り上げる（上方丸め）
Public Function RoundUpTime(ByVal hours As Double) As Double
    RoundUpTime = Application.WorksheetFunction.Ceiling(hours, ROUND_UNIT_HOURS)
End Function

''' 出勤時刻と退勤時刻から勤務時間を計算する
''' 昼休憩（1時間）を差し引きます
Public Function CalcWorkHours(ByVal timeIn As Date, ByVal timeOut As Date) As Double
    Dim startHours As Double, endHours As Double
    startHours = RoundUpOnWork(timeIn) '0602 fix
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
        CalcOvertimeHours = RoundUpTime(workHours - 0.5 - STANDARD_WORK_HOURS) '0.5引いてから残業
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
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(Trim(CStr(ws.Cells(i, 1).value)), key, vbTextCompare) = 0 Then
            GetSetting = Trim(CStr(ws.Cells(i, 2).value))
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
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = 1 To lastRow
        If StrComp(Trim(CStr(ws.Cells(i, 1).value)), key, vbTextCompare) = 0 Then
            ws.Cells(i, 2).value = value
            Exit Sub
        End If
    Next i
    
    ' キーが見つからない場合、末尾に追加
    ws.Cells(lastRow + 1, 1).value = key
    ws.Cells(lastRow + 1, 2).value = value
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
        .title = title
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
        .title = title
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
Public Function GetExcelFiles_廃止(ByVal folderPath As String) As Variant
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
       ' GetExcelFiles = Array()
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
    
    If fileList.count = 0 Then
        'GetExcelFiles = Array()
    Else
       ' GetExcelFiles = fileList.ToArray()
    End If
End Function

' 指定フォルダ内のExcelファイル一覧（xls/xlsx/xlsm）を取得する
' 戻り値: ファイルパスの配列
Public Function GetExcelFiles(ByVal folderPath As String) As Variant
    ' フォルダが存在しない場合、空配列を返す
    Dim normalizedPath As String
    normalizedPath = Trim(folderPath)
    If Len(normalizedPath) > 0 And (Right$(normalizedPath, 1) = "\" Or Right$(normalizedPath, 1) = "/") Then
        normalizedPath = Left$(normalizedPath, Len(normalizedPath) - 1)
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(normalizedPath) Then
        GetExcelFiles = Array()
        Exit Function
    End If
    
    Dim fileList As Collection
    Set fileList = New Collection
    
    Dim folderObj As Object
    Set folderObj = fso.GetFolder(normalizedPath)
    
    Dim fileObj As Object
    For Each fileObj In folderObj.files
        Dim ext As String
        ext = LCase$(fso.GetExtensionName(fileObj.Name))
        
        Select Case ext
            Case "xls", "xlsx", "xlsm"
                ' Dir 関数では環境依存で文字化けすることがあるため、
                ' Unicode を保持できる FSO の Path をそのまま保持する
                fileList.Add CStr(fileObj.Path)
        End Select
    Next fileObj
    
    If fileList.count = 0 Then
        GetExcelFiles = Array()
    Else
        Dim resultArray() As String
        ReDim resultArray(0 To fileList.count - 1)
        Dim idx As Long
        For idx = 1 To fileList.count
            resultArray(idx - 1) = fileList(idx)
        Next idx
        GetExcelFiles = resultArray
    End If
End Function

' Collectionを配列に変換するヘルパー関数
Private Function CollectionToArray(col As Collection) As Variant
    Dim arr() As String
    Dim i As Integer
    
    If col.count = 0 Then
        CollectionToArray = Array()
        Exit Function
    End If
    
    ReDim arr(1 To col.count)
    For i = 1 To col.count
        arr(i) = col(i)
    Next i
    CollectionToArray = arr
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

''' 列文字（A?Z）を列番号（1-based）に変換
Public Function ColLetterToNum(ByVal colLetter As String) As Long
    ColLetterToNum = Range(colLetter & "1").Column
End Function

'''ログ一覧
Sub LogToListBox(msg As String)
    Debug.Print msg
    frmLog.ListBox1.AddItem Format(Now(), "HH:MM:SS") & " " & msg
    frmLog.ListBox1.TopIndex = frmLog.ListBox1.ListCount - 1
End Sub
'''日付転換
Function ToDate(str As String) As String
    Dim datePart As String
    datePart = Left(Trim(str), 10)
    ToDate = Replace(Replace(datePart, "-", ""), "/", "")
End Function


Function ConvertToDateStr(ByVal strDate As Variant) As String
    Dim tempStr As String
    
    
    tempStr = Trim(CStr(strDate))
    
  
    If Len(tempStr) <> 8 Or Not IsNumeric(tempStr) Then
        ConvertToDateStr = ""
        Exit Function
    End If
    
    ' yyyy/mm/dd
    ConvertToDateStr = Left(tempStr, 4) & "/" & Mid(tempStr, 5, 2) & "/" & Right(tempStr, 2)
End Function

' ==============================================
' 機能：8桁の有効な日付文字列かを判定（Boolean返却）
' 判定ルール：
'   1. 文字列長が8桁であること
'   2. 数値のみで構成されていること
'   3. 実在する年月日であること（日付の繰り上げを禁止）
' 返却値：正常 → True、異常 → False
' ==============================================
Function IsValid8Date(ByVal val As Variant) As Boolean
    Dim str As String
    Dim y As Integer, m As Integer, d As Integer
    Dim checkDate As Date
    
    ' 初期値をFalseに設定
    IsValid8Date = False
    
    ' 文字列に変換し、前後の空白を除去
    str = Trim(CStr(val))
    
    ' 1. 8桁チェック
    If Len(str) <> 8 Then Exit Function
    
    ' 2. 数値チェック
    If Not IsNumeric(str) Then Exit Function
    
    ' 年月日を数値に分解
    y = CInt(Left(str, 4))
    m = CInt(Mid(str, 5, 2))
    d = CInt(Right(str, 2))
    
    ' 月の範囲チェック (1～12)
    If m < 1 Or m > 12 Then Exit Function
    
    ' エラー制御を開始
    On Error Resume Next
    
    ' 日付に変換
    checkDate = DateSerial(y, m, d)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    
    ' 【最重要チェック】
    ' 変換後の日付が、元の年月日と完全に一致するか確認
    If Year(checkDate) = y And month(checkDate) = m And day(checkDate) = d Then
        IsValid8Date = True
    End If
End Function


Sub openTool()
  frmMain.Show
End Sub

' 引数：氏名、対象月、対象日、対象年 → 該当の作業内容を返却
Function GetWorkContentByDate( _
    ByVal personName As String, _
    ByVal month As Long, _
    ByVal day As Long, _
    ByVal targetY As Long) As String
    
    Dim personData As String
    Dim rowsData() As String
    Dim rowItem As Variant
    Dim fields() As String
    Dim currentDate As Date
    Dim startDate As Date, endDate As String
    
    ' 設定データ読み込み
    personData = GetSetting("personData")
    rowsData = Split(personData, vbCrLf)
    
    ' 判定用の日付作成
    currentDate = DateSerial(targetY, month, day)
    
    ' データを1行ずつループ
    For Each rowItem In rowsData
        If Trim(rowItem) = "" Then GoTo NextLine
        
        ' Tabで列を分割
        fields = Split(rowItem, vbTab)
        
        ' 列数が足りているかチェック
        If UBound(fields) >= 5 Then
            
            ' ===== ここが重要：氏名が一致するか確認 =====
            Dim dataName As String
            dataName = Trim(fields(0)) ' 1列目：氏名
            
            If dataName = personName Then
                
                ' 日付区間を取得
                Dim dateRange As String
                dateRange = Trim(fields(1)) ' 2列目：期間
                Dim dateParts() As String
                dateParts = Split(dateRange, "-")
                
                If UBound(dateParts) = 1 Then
                    ' 開始日と終了日を作成（yyyy/mm/dd-yyyy/mm/dd 形式対応）
                    Dim sParts() As String, eParts() As String
                    Dim sY As Long, sM As Long, sD As Long
                    Dim eY As Long, eM As Long, eD As Long
                    sParts = Split(Trim(dateParts(0)), "/")
                    eParts = Split(Trim(dateParts(1)), "/")
                    Dim parsedOK As Boolean
                    parsedOK = True
                    If UBound(sParts) >= 2 Then
                        sY = CLng(sParts(0)) : sM = CLng(sParts(1)) : sD = CLng(sParts(2))
                    ElseIf UBound(sParts) = 1 Then
                        sY = targetY : sM = CLng(sParts(0)) : sD = CLng(sParts(1))
                    Else
                        parsedOK = False
                    End If
                    If parsedOK Then
                        If UBound(eParts) >= 2 Then
                            eY = CLng(eParts(0)) : eM = CLng(eParts(1)) : eD = CLng(eParts(2))
                        ElseIf UBound(eParts) = 1 Then
                            eY = targetY : eM = CLng(eParts(0)) : eD = CLng(eParts(1))
                        Else
                            parsedOK = False
                        End If
                    End If
                    If parsedOK Then
                        startDate = DateSerial(sY, sM, sD)
                        endDate = DateSerial(eY, eM, eD)
                        
                        ' 日付が範囲内の場合、作業内容を返す
                        If currentDate >= startDate And currentDate <= endDate Then
                            GetWorkContentByDate = Trim(fields(2))
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
NextLine:
    Next rowItem
    
    ' 該当データなし
    GetWorkContentByDate = ""
End Function
