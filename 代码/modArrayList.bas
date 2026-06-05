Attribute VB_Name = "modArrayList"
' CreateObject("System.Collections.ArrayList")  使用不可のため、共通
'=============================================
' 共用：ArrayList 代替関数（32bit/64bit 両対応）
' オートメーションエラー完全解消
'=============================================
Public Function NewArrayList() As Collection
    Set NewArrayList = New Collection
End Function

'=============================================
' Collection を配列に変換（旧ToArray() と同じ動作）
'=============================================
Public Function ToArray(col As Variant) As Variant
    If TypeName(col) <> "Collection" Then
        ToArray = Array()
        Exit Function
    End If
    
    If col.count = 0 Then
        ToArray = Array()
        Exit Function
    End If
    
    Dim arr() As String
    ReDim arr(1 To col.count)
    
    Dim i As Long
    For i = 1 To col.count
        arr(i) = col(i)
    Next i
    
    ToArray = arr
End Function

'=============================================
' Collection に要素を追加（旧Add() と同じ）
'=============================================
Public Sub AddToList(ByRef list As Variant, val As Variant)
    If TypeName(list) = "Collection" Then
        list.Add val
    End If
End Sub
