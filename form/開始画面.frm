VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 開始画面 
   Caption         =   "操作開始"
   ClientHeight    =   3075
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "開始画面.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "開始画面"
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
        MsgBox "担当者を入力して下さい。"
    Else
        Unload Me
        組合員名簿管理.Show
    End If
    
End Sub

Private Sub CommandButton3_Click()

    If ComboBox1.Value = "" Then
        MsgBox "担当者を入力して下さい。"
    Else
        Unload Me
        内容入力.Show
    End If

End Sub

Private Sub CommandButton4_Click()

Unload Me

End Sub

Private Sub UserForm_initialize()
    Dim sht     As Worksheet                '// 参照シート
    Set sht = ThisWorkbook.Worksheets("Master")

    With ComboBox1 'リストボックスの表示'
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

