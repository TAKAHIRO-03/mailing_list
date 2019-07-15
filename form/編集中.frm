VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 編集中 
   Caption         =   "編集中"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5280
   OleObjectBlob   =   "編集中.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "編集中"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Unload 編集中

End Sub

Private Sub CommandButton2_Click()

If all_address.Exists(TextBox2.Text) Then

MsgBox "名簿に居る人と同じ名前にすることは出来ません。"

ElseIf TextBox1.Text = "" Then

MsgBox "部署が空白です。"

ElseIf TextBox2.Text = "" Then

MsgBox "名前が空白です。"
 
ElseIf TextBox3.Text = "" Then

MsgBox "メールアドレスが空白です。"

Else
    busho = TextBox1.Text
    name_edit = TextBox2.Text
    mead_edit = TextBox3.Text
    Unload 編集中
End If

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_initialize()

Set all_address = a.all_addresses
all_address.Remove (name_edit)

TextBox1 = busho
TextBox2 = name_edit
TextBox3 = mead_edit

End Sub
