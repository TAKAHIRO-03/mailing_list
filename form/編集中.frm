VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �ҏW�� 
   Caption         =   "�ҏW��"
   ClientHeight    =   2940
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5280
   OleObjectBlob   =   "�ҏW��.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�ҏW��"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

Unload �ҏW��

End Sub

Private Sub CommandButton2_Click()

If all_address.Exists(TextBox2.Text) Then

MsgBox "����ɋ���l�Ɠ������O�ɂ��邱�Ƃ͏o���܂���B"

ElseIf TextBox1.Text = "" Then

MsgBox "�������󔒂ł��B"

ElseIf TextBox2.Text = "" Then

MsgBox "���O���󔒂ł��B"
 
ElseIf TextBox3.Text = "" Then

MsgBox "���[���A�h���X���󔒂ł��B"

Else
    busho = TextBox1.Text
    name_edit = TextBox2.Text
    mead_edit = TextBox3.Text
    Unload �ҏW��
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
