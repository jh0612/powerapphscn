VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLog 
   Caption         =   "ログ一覧"
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   OleObjectBlob   =   "frmLog.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyData As New MSForms.DataObject
Private Sub btnBack_Click()
 Me.Hide
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.ListBox1.listIndex = -1 Then Exit Sub
    

    Dim s As String
    s = Me.ListBox1.value
    

    MyData.SetText s
    MyData.PutInClipboard
    
    MsgBox "コピーしました" & s
End Sub
