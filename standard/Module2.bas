Attribute VB_Name = "Module2"
'�������t�H�[����\�����Ă���ۂɎ��s����
'���C�������v���O����
Option Explicit
Public flag As Integer

Sub main()

flag = 1
Dim i As Long
Application.ScreenUpdating = False

For i = 0 To 100000
    If flag = 1 Then
    'OS�ɏ�����Ԃ�(��ʕ`����X�V)
    DoEvents
    ������.Label1.Caption = "�������ł�..." & i & "��"
    ElseIf flag = 0 Then
        Exit For
    End If
Next i

End Sub


