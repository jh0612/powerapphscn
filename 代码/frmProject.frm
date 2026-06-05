VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProject 
   Caption         =   "プロジェクト基本情報・テンプレート設定 "
   ClientHeight    =   10410
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13755
   OleObjectBlob   =   "frmProject.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "frmProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==================================================
' frmProject コード
'==================================================

Private Sub UserForm_Initialize()
    'cboServiceMode.AddItem SERVICE_MODE_TM
   ' cboServiceMode.AddItem SERVICE_MODE_LUMPSUM
    
    'cboServiceMode
     Dim arr, i As Long
    Dim sht As Worksheet
    Set sht = Sheets("name")
    
    arr = sht.Range("I2", sht.Cells(sht.Rows.count, "I").End(xlUp)).value
    
    cboServiceMode.Clear 'クリア
       For i = LBound(arr) To UBound(arr)
        If Trim(arr(i, 1)) <> "" Then
            cboServiceMode.AddItem arr(i, 1)
        End If
    Next i
    
    txtProjectName.Text = GetSetting(SET_PROJECT_NAME)
    txtSupplier.Text = GetSetting(SET_SUPPLIER)
    txtServiceContent.Text = GetSetting(SET_SERVICE_CONTENT)
    'txtDeliverable.Text = GetSetting(SET_DELIVERABLE)
    txtMilestone.Text = GetSetting(SET_MILESTONE)
    'txtContractNo.Text = GetSetting(SET_CONTRACT_NO)
    'txtPO.Text = GetSetting(SET_PO)
    'txtOrderNo.Text = GetSetting(SET_ORDER_NO)
    
    Dim svcMode As String
    svcMode = GetSetting(SET_SERVICE_MODE)
    If Len(svcMode) > 0 Then cboServiceMode.Text = svcMode Else cboServiceMode.listIndex = 0
    
    txtHSCNPath.Text = GetSetting(SET_HSCN_PATH)
    txtCompanyPath.Text = GetSetting(SET_COMPANY_PATH)
    txtOrderPath.Text = GetSetting(SET_ORDER_PATH)
    txtOutputPath.Text = GetSetting(SET_OUTPUT_PATH)
End Sub

Private Sub btnHSCNBrowse_Click()
    Dim fp As String: fp = BrowseFile("HSCN?收?テンプレートを選択")
    If Len(fp) > 0 Then txtHSCNPath.Text = fp
End Sub

Private Sub btnCompanyBrowse_Click()
    Dim fp As String: fp = BrowseFile("会社明細テンプレートを選択")
    If Len(fp) > 0 Then txtCompanyPath.Text = fp
End Sub

Private Sub btnOrderBrowse_Click()
    Dim fp As String: fp = BrowseFile("???价兼?注?テンプレートを選択")
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
    'SaveSetting SET_DELIVERABLE, txtDeliverable.Text
    SaveSetting SET_MILESTONE, txtMilestone.Text
    'SaveSetting SET_CONTRACT_NO, txtContractNo.Text
    'SaveSetting SET_PO, txtPO.Text
    'SaveSetting SET_ORDER_NO, txtOrderNo.Text
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
