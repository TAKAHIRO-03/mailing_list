VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ���e���� 
   Caption         =   "���e����"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7215
   OleObjectBlob   =   "���e����.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "���e����"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

 Content = Split(TextBox3.Text, vbCrLf)
 Title = TextBox2.Text
 
 If UBound(Content) = 0 Then
    MsgBox "���e����͂��ĉ������B"
    Exit Sub
 ElseIf Title = "" Then
    MsgBox "�^�C�g������͂��ĉ������B"
    Exit Sub
  End If
 
 Unload Me
 �ŏI�m�F.Show
 
End Sub
Private Sub CommandButton2_Click()
 Unload Me
End Sub


Private Sub Label1_Click()

End Sub

Private Sub UserForm_Terminate()
    Set g = Nothing
    Set name_busho_meado = Nothing
End Sub

Private Sub UserForm_initialize()
    Call GETINFO_CLASS
    Dim list_name() As String
    Dim i As Integer
    Dim j As Integer
    Set name_busho_meado = CreateObject("Scripting.Dictionary")
    Set name_busho_meado = g.name_busho_meados
    Dim curKey As Variant

    With ListBox1
        .Clear
          For Each curKey In name_busho_meado
                list_name = name_busho_meado.Item(curKey)
                .AddItem ""
                .List(.ListCount - 1, 0) = curKey
                .List(.ListCount - 1, 1) = list_name(1, 2)
          Next
    End With
    
    Label1.Caption = Person
    Label8.Caption = name_busho_meado.Count & "��"
    
End Sub



