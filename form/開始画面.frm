VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �J�n��� 
   Caption         =   "����J�n"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "�J�n���.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�J�n���"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ComboBox1_Change()

    With ComboBox1
        Person = .List(.ListIndex)
    End With
    
End Sub

Private Sub CommandButton2_Click()

    If ComboBox1.Value = "" Then
        MsgBox "�S���҂���͂��ĉ������B"
    Else
        Unload Me
        �g��������Ǘ�.Show
    End If
    
End Sub

Private Sub CommandButton3_Click()

    If ComboBox1.Value = "" Then
        MsgBox "�S���҂���͂��ĉ������B"
    Else
        Unload Me
        ���e����.Show
    End If

End Sub

Private Sub CommandButton4_Click()

Unload Me

End Sub

Private Sub UserForm_initialize()
    Dim sht     As Worksheet                '// �Q�ƃV�[�g
    Set sht = ThisWorkbook.Worksheets("Master")

    With ComboBox1 '���X�g�{�b�N�X�̕\��'
        .AddItem sht.Cells(2, 2)
        .AddItem sht.Cells(3, 2)
        .AddItem sht.Cells(4, 2)
        .AddItem sht.Cells(5, 2)
        .AddItem sht.Cells(6, 2)
        .AddItem sht.Cells(7, 2)
        .AddItem sht.Cells(8, 2)
        .AddItem sht.Cells(9, 2)
        .AddItem sht.Cells(10, 2)
        .AddItem sht.Cells(11, 2)
    End With

End Sub

