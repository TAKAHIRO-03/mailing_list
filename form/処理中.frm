VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ������ 
   Caption         =   "������"
   ClientHeight    =   2265
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "������.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "������"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub Show_UserForm1()

'�������t�H�[���̕\��(���[�_��)
������.Show vbModal

End Sub

Private Sub UserForm_Activate()
'�}�E�X�|�C���^�[�������v�ɕύX
Application.Cursor = xlWait
'���C���̏������Ăяo��
Call GETMAIL_CLASS
Call m.submit_mail(name_busho_meado, Content, Title)
'�}�E�X�|�C���^�[�����ɖ߂�
Application.Cursor = xlDefault
'�����I���̃��b�Z�[�W��\��
������.Label1.Caption = "�������I�����܂����B"

End Sub

'����{�^���������ꂽ�ۂ̏���
Private Sub CommandButton1_Click()
'�t�H�[�����A�����[�h����
flag = 0
Unload ������

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'�t�H�[���̕���{�^���̖�����
If CloseMode = 0 Then Cancel = True

End Sub


