VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPerson 
   Caption         =   "作業人員設定"
   ClientHeight    =   9420.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13080
   OleObjectBlob   =   "frmPerson.frx":0000
   StartUpPosition =   2  '画面の中央
End
Attribute VB_Name = "frmPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'==================================================
' frmPerson コード
'==================================================

Private m_editIndex As Long

Private Sub UserForm_Initialize()
    m_editIndex = -1
        'Chief,Architect/Designer,PM,Senior 2,Senior 1,Junior,Beginner
    cboLevel.AddItem "Chief"
    cboLevel.AddItem "Architect/Designer"
    cboLevel.AddItem "PM"
    cboLevel.AddItem "Senior 1"
    cboLevel.AddItem "Senior 2"
    cboLevel.AddItem "Junior"
    cboLevel.AddItem "Beginner"
    cboLevel.listIndex = 0

    
    
     Dim arr, i As Long
    Dim sht As Worksheet
    Set sht = Sheets("name")
    
    arr = sht.Range("A2", sht.Cells(sht.Rows.count, "A").End(xlUp)).value
    
    cboName.Clear 'クリア
       For i = LBound(arr) To UBound(arr)
        If Trim(arr(i, 1)) <> "" Then
            cboName.AddItem arr(i, 1)
        End If
    Next i
    '会社
     arr = sht.Range("E2", sht.Cells(sht.Rows.count, "E").End(xlUp)).value
    cboCompany.Clear 'クリア
      For i = LBound(arr) To UBound(arr)
        If Trim(arr(i, 1)) <> "" Then
            cboCompany.AddItem arr(i, 1)
        End If
    Next i

    
    Call LoadPersonList
End Sub

Private Sub LoadPersonList()
    lstPersonList.Clear
    ' lstPersonList.AddItem "番号"
               ' lstPersonList.list(0, 1) = "氏名"
                'lstPersonList.list(0, 2) = "作業期間"
                'lstPersonList.list(0, 3) = "作業内容"
                'lstPersonList.list(0, 4) = "会社名"
                'lstPersonList.list(0, 5) = "作業番号"
                'lstPersonList.list(0, 6) = "レベル"
                
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
                lstPersonList.list(rowIdx, 1) = cols(0)
                lstPersonList.list(rowIdx, 2) = cols(1)
                lstPersonList.list(rowIdx, 3) = cols(2)
                lstPersonList.list(rowIdx, 4) = cols(3)
                lstPersonList.list(rowIdx, 5) = cols(4)
                lstPersonList.list(rowIdx, 6) = cols(5)
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
    lstPersonList.list(rowIdx, 1) = cboName.Text
    lstPersonList.list(rowIdx, 2) = txtWorkPeriod.Text
     lstPersonList.list(rowIdx, 3) = txtWorkContent.Text
     lstPersonList.list(rowIdx, 4) = cboCompany.Text
    lstPersonList.list(rowIdx, 5) = txtWorkNo.Text
    lstPersonList.list(rowIdx, 6) = cboLevel.Text
    Call ClearInputs
End Sub

Private Sub lstPersonList_Click()
    If lstPersonList.listIndex < 0 Then Exit Sub
    m_editIndex = lstPersonList.listIndex
    cboName.Text = lstPersonList.list(m_editIndex, 1)
    txtWorkPeriod.Text = lstPersonList.list(m_editIndex, 2)
    txtWorkContent.Text = lstPersonList.list(m_editIndex, 3)
    cboCompany.Text = lstPersonList.list(m_editIndex, 4)
    txtWorkNo.Text = lstPersonList.list(m_editIndex, 5)
    Dim lvl As String: lvl = lstPersonList.list(m_editIndex, 6)
    Dim i As Long
    For i = 0 To cboLevel.ListCount - 1
        If cboLevel.list(i) = lvl Then cboLevel.listIndex = i: Exit For
    Next i
End Sub

Private Sub btnUpdate_Click()
    If m_editIndex < 0 Then
        MsgBox "更新する行を選択してください。", vbExclamation: Exit Sub
    End If
    If Len(Trim(cboName.Text)) = 0 Then
        MsgBox "氏名を入力してください。", vbExclamation: Exit Sub
    End If
    lstPersonList.list(m_editIndex, 1) = cboName.Text
    lstPersonList.list(m_editIndex, 2) = txtWorkPeriod.Text
        lstPersonList.list(m_editIndex, 3) = txtWorkContent.Text
    lstPersonList.list(m_editIndex, 4) = cboCompany.Text
    lstPersonList.list(m_editIndex, 5) = txtWorkNo.Text
    lstPersonList.list(m_editIndex, 6) = cboLevel.Text
    Call ClearInputs: m_editIndex = -1
End Sub

Private Sub btnDelete_Click()
    If lstPersonList.listIndex < 0 Then
        MsgBox "削除する行を選択してください。", vbExclamation: Exit Sub
    End If
    If MsgBox("選択行を削除しますか？", vbYesNo + vbQuestion, "削除確認") <> vbYes Then Exit Sub
    lstPersonList.RemoveItem lstPersonList.listIndex
    Dim i As Long
    For i = 0 To lstPersonList.ListCount - 1
        lstPersonList.list(i, 0) = CStr(i + 1)
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
        If Len(lstPersonList.list(i, 1)) > 0 Then
            Dim line As String
            line = lstPersonList.list(i, 1) & vbTab & _
                   lstPersonList.list(i, 2) & vbTab & _
                   lstPersonList.list(i, 3) & vbTab & _
                   lstPersonList.list(i, 4) & vbTab & _
                   lstPersonList.list(i, 5) & vbTab & _
                   lstPersonList.list(i, 6)
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
    If cboLevel.ListCount > 0 Then cboLevel.listIndex = 0
End Sub

