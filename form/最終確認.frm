VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} �ŏI�m�F 
   Caption         =   "�ŏI�m�F"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7575
   OleObjectBlob   =   "�ŏI�m�F.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "�ŏI�m�F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim rc As Integer
    rc = MsgBox("�{���Ƀ��[���𑗐M���Ă���낵���ł��傤���B", vbYesNo + vbQuestion, "�m�F")
    
    If rc = vbNo Then
          Exit Sub
    ElseIf rc = vbYes Then
         ������.Show
    End If
    Unload Me
 
End Sub

Private Sub CommandButton2_Click()

Unload Me

End Sub

Private Sub UserForm_Terminate()
    Set g = Nothing
    Set m = Nothing
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
       
    With ListBox3
        .Clear
          For Each curKey In name_busho_meado
                list_name = name_busho_meado.Item(curKey)
                .AddItem ""
                .List(.ListCount - 1, 0) = curKey
                .List(.ListCount - 1, 1) = list_name(1, 2)
          Next
    End With
    
    Label8.Caption = Person
    Label7.Caption = Title
    Label11.Caption = name_busho_meado.Count & "��"

    With ListBox4
            .AddItem ""
            .AddItem "�����@�������œ��͂���܂��B"
            .AddItem ""
            .AddItem "�����@�������œ��͂���܂��B"
            .AddItem ""
        For j = 0 To UBound(Content)
            .AddItem Content(j)
        Next j
    End With
    
End Sub




