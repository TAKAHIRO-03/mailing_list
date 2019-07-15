VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 処理中 
   Caption         =   "処理中"
   ClientHeight    =   2265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "処理中.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "処理中"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Show_UserForm1()

'処理中フォームの表示(モーダル)
処理中.Show vbModal

End Sub

Private Sub UserForm_Activate()
'マウスポインターを砂時計に変更
Application.Cursor = xlWait
'メインの処理を呼び出す
Call GETMAIL_CLASS
Call m.submit_mail(name_busho_meado, Content, Title)
'マウスポインターを元に戻す
Application.Cursor = xlDefault
'処理終了のメッセージを表示
処理中.Label1.Caption = "処理が終了しました。"

End Sub

'閉じるボタンが押された際の処理
Private Sub CommandButton1_Click()
'フォームをアンロードする
flag = 0
Unload 処理中

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'フォームの閉じるボタンの無効化
If CloseMode = 0 Then Cancel = True

End Sub


