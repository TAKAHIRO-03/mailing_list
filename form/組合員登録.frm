VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 組合員登録 
   Caption         =   "組合員登録"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5820
   OleObjectBlob   =   "組合員登録.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "組合員登録"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()

    With ComboBox1
        busho = .List(.ListIndex)
    End With

End Sub

Private Sub CommandButton1_Click()
    Dim busho_add As String
    Dim name_add As String
    Dim mead_add As String
    Dim sht As Worksheet  'マスターシートを参照する。
    Dim busho_last As Range '列の一番下に代入する用
    Dim busho_ad As String 'アドレス変換用
    Dim r_row As Long '代入時の行を取得
    Dim r_column As Long '代入時の列を取得
    Dim last_r As Long
    Dim last_c As Long
    Dim last_cell As Long
    
    Set all_address = a.all_addresses
    Set sht = a.Mysht  'マスターシートを参照する。
    CommandButton1.Enabled = False 'ボタン連打防止
    busho_add = busho
    name_add = TextBox1.Text
    mead_add = TextBox2.Text & "@rike-vita.co.jp"
    Set busho_all_address = a.busho_all_address
    list_busho = busho_all_address.keys
    list_address = busho_all_address.items
       
    If busho = "" Then
          MsgBox "部署を選んで下さい。"
          CommandButton1.Enabled = True
          Exit Sub
    ElseIf name_add = "" Then
          MsgBox "名前が未入力です。"
          CommandButton1.Enabled = True
          Exit Sub
    ElseIf TextBox2.Text = "" Then
          MsgBox "メアドが未入力です。"
          CommandButton1.Enabled = True
          Exit Sub
    ElseIf all_address.Exists(name_add) Then
          MsgBox "名簿に既に登録されています。"
          CommandButton1.Enabled = True
          Exit Sub
    End If
    
    busho_ad = busho_all_address(busho)
    last_c = sht.Range(busho_ad).Column
    last_cell = sht.Cells(Rows.Count, last_c).End(xlUp).Row + 1
    r_column = sht.Range(busho_ad).Column
    sht.Cells(last_cell, r_column) = busho
    sht.Cells(last_cell, r_column + 1) = name_add
    sht.Cells(last_cell, r_column + 2) = mead_add
          
    Set sht = Nothing
    Set name_busho_meado = Nothing
    Set all_address = Nothing
    CommandButton1.Enabled = True  'ボタン連打防止
    
    TextBox1 = ""
    TextBox2 = ""
    
    MsgBox "登録が完了しました。"
    
End Sub

Private Sub CommandButton3_Click()

  Unload Me
  組合員名簿管理.Show

End Sub

Private Sub UserForm_Terminate()

    Set g = Nothing
    Set a = Nothing
    Set name_busho_meado = Nothing
    Set all_address = Nothing
    Set busho_all_address = Nothing
    
End Sub
Private Sub UserForm_initialize()

    Call GETADD_DELETE
    Dim i As Integer
    Set all_address = a.all_addresses
    Set busho_all_address = a.busho_all_address
    list_busho = busho_all_address.keys
    
    With ComboBox1 'リストボックスの表示'
        For i = 0 To UBound(list_busho)
            .AddItem list_busho(i)
        Next i
    End With
        
End Sub

