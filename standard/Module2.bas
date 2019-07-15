Attribute VB_Name = "Module2"
'処理中フォームを表示している際に実行する
'メイン処理プログラム
Option Explicit
Public flag As Integer

Sub main()

flag = 1
Dim i As Long
Application.ScreenUpdating = False

For i = 0 To 100000
    If flag = 1 Then
    'OSに処理を返す(画面描画を更新)
    DoEvents
    処理中.Label1.Caption = "処理中です..." & i & "件"
    ElseIf flag = 0 Then
        Exit For
    End If
Next i

End Sub


