VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMain 
   Caption         =   "勤怠データ集計・請求書自動生成ツール"
   ClientHeight    =   11820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14160
   OleObjectBlob   =   "frmMain.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==================================================
' frmMain コード
' メイン画面のイベント処理
'==================================================

' フルパスを保持するコレクション（モジュールレベル変数）
Private m_filePathsList As Object
' Dictionary: Key=ファイル名, Value=フルパス

Private Sub btnLog_Click()
frmLog.Show vbModal
End Sub

Private Sub btnThisMonth_Click()
    txtPeriod.Text = Format(DateSerial(Year(Date), month(Date), 1), "yyyymmdd")
    txtPeriodTo.Text = Format(DateSerial(Year(Date), month(Date) + 1, 0), "yyyymmdd")
End Sub

'==================================================
' frmMain コード
' メイン画面のイベント処理
'==================================================

'--------------------------------------------------
' フォーム初期化
'--------------------------------------------------

Private Sub UserForm_Initialize()
    ' フルパスマップの初期化
    Set m_filePathsList = CreateObject("Scripting.Dictionary")
    
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
            If cboLocation.list(i) = lastLocation Then
                cboLocation.listIndex = i
                Exit For
            End If
        Next i
    Else
        cboLocation.listIndex = 0
    End If
    
    ' フォルダパスを復元
    txtFolderPath.Text = GetSetting("LastFolderPath")
    
    ' 担当者コンボボックスに選択肢を追加
    cboPerson.AddItem "(すべて)"
    cboPerson.listIndex = 0
    
    ' ステータス初期表示
    lblStatus.Caption = "待機中"
    lblCount.Caption = "選択件数：0 件"
    
    ' ファイル一覧を更新
    'If Len(txtFolderPath.Text) > 0 Then
    '    Call RefreshFileList
    'End If
        ' 業務期間を当月の初日と末日で自動設定（v3改善）
        
    Dim firstDay As String
    Dim lastDay As String
    firstDay = Format(DateSerial(Year(Date), month(Date), 1), "yyyymmdd")
    lastDay = Format(DateSerial(Year(Date), month(Date) + 1, 0), "yyyymmdd")
    
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
        'Call RefreshFileList
    End If
End Sub

'--------------------------------------------------
' フォルダパス変更時
'--------------------------------------------------

Private Sub txtFolderPath_Change()
    'Call RefreshFileList
    Call SafeRefreshFileList
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
    On Error GoTo ErrHandler
    
    If cboLocation.listIndex < 0 Then Exit Sub
    
    ' 保存最后選択的据点
    SaveSetting "LastLocation", cboLocation.Text
    
    ' 自動設定据点フォルダ路径
    Dim autoPath As String
    autoPath = ThisWorkbook.Path & "\" & cboLocation.Text
    
    ' ?当?文件?存在?自?填入路径（否?保留手?路径）
    If Len(Dir(autoPath, vbDirectory)) > 0 Then
        txtFolderPath.Text = autoPath
        SaveSetting "LastFolderPath", autoPath
    End If
    
    ' 刷新担当者和文件列表
    Call RefreshPersonList
    Call SafeRefreshFileList
    'Call RefreshFileList
    Exit Sub

ErrHandler:
    '　エラーが発生なら
    lstFiles.Clear
    lblCount.Caption = "選択件数：0 件（エラーが発生　cboLocation_Change()）"
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
' 全?チェックボックス
'--------------------------------------------------

Private Sub chkSelectAll_Click()
    Dim i As Long
    For i = 0 To lstFiles.ListCount - 1
        lstFiles.Selected(i) = chkSelectAll.value
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
    Dim selectedFiles As Collection
    Set selectedFiles = New Collection
    
    ' FSO でファイルパスを正確に取得するための準備
    Dim folderPath As String
    folderPath = Trim(txtFolderPath.Text)
    If Len(folderPath) = 0 Then
        MsgBox "フォルダパスを選択してください。", vbExclamation, "確認"
        Exit Sub
    End If
    
    ' フォルダパスの正規化（末尾の\を削除）
    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        MsgBox "フォルダが見つかりません：" & folderPath, vbExclamation, "エラー"
        Exit Sub
    End If
    
    ' 選択されたファイル名から完全パスを構築
    Dim i As Long
    'For i = 0 To lstFiles.ListCount - 1
    '    If lstFiles.Selected(i) Then
    '        Dim fileName As String
    '        fileName = CStr(lstFiles.list(i, 0))
    '
            ' FSO を使用して完全パスを構築（Unicode対応）
    '        If m_filePathsMap.Exists(fileName) Then
    '            Dim fullPath As String
    '            fullPath = CStr(m_filePathsMap(fileName))
    '            selectedFiles.Add fullPath
    '            Debug.Print "Selected: " & fileName & " -> " & fullPath
    '        Else
    '            Debug.Print "File not found: " & fullPath
    '        End If
    '    End If
    'Next i
    
    
    
    For i = 0 To lstFiles.ListCount - 1
        If lstFiles.Selected(i) Then
            ' Dictionaryから対応するフルパスを取得（インデックスをキーとして使用）
            Dim indexKey As String
            indexKey = CStr(i)
            
            If m_filePathsList.Exists(indexKey) Then
                Dim fullPath As String
                fullPath = CStr(m_filePathsList(indexKey))
                
                Debug.Print "Selected index: " & i & ", Path: " & fullPath
                
                If Len(fullPath) > 0 Then
                    selectedFiles.Add fullPath
                Else
                    Debug.Print "Warning: Empty path for index " & i
                End If
            Else
                Debug.Print "Warning: Index " & i & " not in map"
            End If
        End If
    Next i
    
    
    If selectedFiles.count = 0 Then
        MsgBox "集計するファイルを選択してください。", vbExclamation, "確認"
        Exit Sub
    End If
    
    ' 拠点・期間を取得
    Dim location As String
    location = ""
    If cboLocation.listIndex >= 0 Then
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
    ReDim filePaths(0 To selectedFiles.count - 1)
    For i = 1 To selectedFiles.count
        filePaths(i - 1) = CStr(selectedFiles(i))
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
        logMsg = NowString() & " - 集計完了：" & selectedFiles.count & "ファイル → " & _
                 SummaryData.count & "名"
        lstHistory.AddItem logMsg, 0
        
        MsgBox "勤怠集計が完了しました。" & vbCrLf & vbCrLf & _
               "集計ファイル数：" & selectedFiles.count & vbCrLf & _
               "集計人員数：" & SummaryData.count & "名" & vbCrLf & _
               "処理時間：" & Format(elapsed, "0.0") & "秒", vbInformation, "集計完了"
    Else
        lblStatus.Caption = "集計に失敗しました"
        lblStatus.ForeColor = &HFF&
        MsgBox "勤怠集計中にエラーが発生しました。詳細はイミディエイトウィンドウを確認してください。", vbCritical, "エラー"
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
    Set m_filePathsList = CreateObject("Scripting.Dictionary")
    
    Dim folderPath As String
    folderPath = Trim(txtFolderPath.Text)
    
    If Len(folderPath) = 0 Then Exit Sub
    
    ' フォルダパスの正規化（末尾の\を削除）
    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        folderPath = Left$(folderPath, Len(folderPath) - 1)
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(folderPath) Then
        chkSelectAll.value = False
        Call UpdateCount
        Exit Sub
    End If
    
    Dim folderObj As Object
    Set folderObj = fso.GetFolder(folderPath)
    
    ' 担当者フィルター
    Dim filterPerson As String
    filterPerson = ""
    If cboPerson.listIndex > 0 Then
        filterPerson = cboPerson.Text
    End If
    
    ' 期間フィルター
    Dim filterPeriod As String
    filterPeriod = Trim(txtPeriod.Text)
    
    Dim fileObj As Object
    For Each fileObj In folderObj.files
        Dim ext As String
        ext = LCase$(fso.GetExtensionName(fileObj.Name))
        
        ' Excelファイルのみ処理
        If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
            Dim fileName As String
            fileName = CStr(fileObj.Name)
            
            Dim fullPath As String
            fullPath = CStr(fileObj.Path)
            
            ' 担当者フィルター（ファイル名に担当者名が含まれるか）
            If Len(filterPerson) > 0 Then
                If InStr(1, fileName, filterPerson, vbTextCompare) = 0 Then
                    GoTo SkipFile
                End If
            End If
            
            ' 期間フィルター
            If Len(filterPeriod) > 0 Then
                If Not IsFileInPeriod(fullPath, filterPeriod) Then
                    GoTo SkipFile
                End If
            End If
            
            Dim listIndex As Long
            listIndex = 0
            
            ' ファイルサイズ
            Dim fileSize As String
            If fileObj.Size < 1024 Then
                fileSize = fileObj.Size & " B"
            ElseIf fileObj.Size < 1048576 Then
                fileSize = Format(fileObj.Size / 1024, "0") & " KB"
            Else
                fileSize = Format(fileObj.Size / 1048576, "0.0") & " MB"
            End If
            
            ' 更新日時
            Dim modifyDate As String
            modifyDate = Format(fileObj.DateLastModified, "yyyy/mm/dd HH:MM")
            
            ' リストに追加（列0=ファイル名、列1=サイズ、列2=更新日時）
            lstFiles.AddItem fileName
            Dim rowIdx As Long
            rowIdx = lstFiles.ListCount - 1
            lstFiles.list(rowIdx, 1) = fileName
            lstFiles.list(rowIdx, 2) = fileSize
            lstFiles.list(rowIdx, 3) = modifyDate
            
            ' Dictionaryにインデックスとフルパスのマッピングを保存
            Dim indexKey As String
            indexKey = CStr(rowIdx)
            m_filePathsList.Add indexKey, fullPath
            
            Debug.Print "Added: Index=" & rowIdx & ", Name=" & fileName & ", Path=" & fullPath
            
            listIndex = listIndex + 1
        End If
        
SkipFile:
    Next fileObj
    
    ' 全?チェックをリセット
    chkSelectAll.value = False
    Call UpdateCount
End Sub
'--------------------------------------------------
' 安全版文件列表刷新（?完整???理）
'--------------------------------------------------
Private Sub SafeRefreshFileList()
    On Error GoTo ErrHandler
    
    lstFiles.Clear
    
    Dim folderPath As String
    folderPath = Trim(txtFolderPath.Text)
    
    ' 路径?空?，?示提示并退出
    If Len(folderPath) = 0 Then
        lblCount.Caption = "選択件数：0 件"
        Exit Sub
    End If
    
    ' 末尾の \ を除去して統一
    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If
    
    ' 文件不存在の場合，友好提示
    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        lblCount.Caption = "選択件数：0 件（フォルダが存在しません）"
        Exit Sub
    End If
    
    ' ?描 .xlsx 文件
    Dim fileName As String
    Dim fileCount As Long
    fileCount = 0
    
    fileName = Dir(folderPath & "\*.xlsx")
    Do While Len(fileName) > 0
        ' 跳?工具文件本身
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
            lstFiles.list(rowIdx, 1) = fileSize
            lstFiles.list(rowIdx, 2) = modDate
            fileCount = fileCount + 1
        End If
        fileName = Dir()
    Loop
    
    ' ?描 .xls 文件（不含 .xlsm 等）
    fileName = Dir(folderPath & "\*.xls")
    Do While Len(fileName) > 0
        ' 跳? .xlsm / .xlsb 等（Dir *.xls 会匹配所有以xls??的?展名）
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
                lstFiles.list(rowIdx, 1) = fileName
                lstFiles.list(rowIdx, 2) = fileSize
                lstFiles.list(rowIdx, 3) = modDate
                fileCount = fileCount + 1
            End If
        End If
        fileName = Dir()
    Loop
    
    ' 更新件数??
    If fileCount = 0 Then
        lblCount.Caption = "選択件数：0 件（ファイルなし）"
    Else
        lblCount.Caption = "選択件数：0 件（合計 " & fileCount & " 件）"
    End If
    
    chkSelectAll.value = False
    Exit Sub

ErrHandler:
    lstFiles.Clear
    lblCount.Caption = "選択件数：0 件（読込エラー）"
End Sub

'--------------------------------------------------
' 担当者一覧更新
'--------------------------------------------------

Private Sub RefreshPersonList()
    ' 現在の選択を保存
    Dim lastPerson As String
    lastPerson = ""
    If cboPerson.listIndex >= 0 Then
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
            If cboPerson.list(i) = lastPerson Then
                cboPerson.listIndex = i
                Exit For
            End If
        Next i
    Else
        cboPerson.listIndex = 0
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



