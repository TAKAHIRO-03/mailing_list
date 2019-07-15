VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 組合員名簿管理 
   Caption         =   "組合員名簿管理"
   ClientHeight    =   6180
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8685
   OleObjectBlob   =   "組合員名簿管理.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "組合員名簿管理"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboBox1_Change()
    Dim list_name() As Variant
    Dim lists() As Variant
    Dim name_ad As String
    Dim busho_ad As String
    Dim i As Integer
    Dim j As Integer
    Dim curKey As Variant
    Dim sht As Worksheet
    Set sht = a.Mysht
    Set all_address = a.all_addresses
    
    With ComboBox1
        busho = .List(.ListIndex)
    End With
        
    With ListBox2
        .Clear
          For Each curKey In all_address
                list_name = all_address.Item(curKey)
                busho_ad = list_name(1, 2)
                If busho = sht.Range(busho_ad) Then
                    .AddItem curKey
                End If
          Next
    End With
     
    Set all_address = Nothing
    Set sht = Nothing
    Erase list_name
End Sub
Private Sub CommandButton1_Click() '追加
    CommandButton1.Enabled = False 'ボタン連打防止
    Dim sht As Worksheet  'マスターシートを参照する。
    Set sht = a.Mysht  'マスターシートを参照する。
    Dim busho_ad As String 'アドレス変換用
    Dim r_row As Long '代入時の行を取得
    Dim r_column As Long '代入時の列を取得
    Dim busho_last As Range '列の一番下に代入する用
    Dim del_name As String '消す用
    Dim del_busho As String '消す用
    Dim del_address As String '消す用
    Dim last_r As Long
    Dim last_c As Long
    Dim last_cell As Long
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim tex As String
    Set name_busho_meado = a.name_busho_meados
    Set all_address = a.all_addresses
    Set busho_all_address = a.busho_all_address
    tex = ListBox3.Text
    list_busho = busho_all_address.keys
    list_address = busho_all_address.items
    
    ReDim busho_address(0 To UBound(list_busho), 1 To 2)
    For j = 0 To UBound(list_busho)
        busho_address(j, 1) = list_busho(j)
        busho_address(j, 2) = list_address(j)
    Next j
    
    If tex = "" Then
          MsgBox "名簿から名前を選んで下さい。"
          CommandButton1.Enabled = True
          Exit Sub
    ElseIf busho = "" Then
          MsgBox "部署を選んで下さい。"
          CommandButton1.Enabled = True
          Exit Sub
    End If
    
    For l = ListBox2.ListCount - 1 To 0 Step -1
            If ListBox2.List(l) = tex Then
               MsgBox "同じ名前の人は追加出来ません。"
               CommandButton1.Enabled = True
               Exit Sub
            End If
    Next l
    

    For i = 0 To UBound(busho_address)
        If busho = busho_address(i, 1) Then
'==================================削除==============================================
           del_name = all_address.Item(tex)(1, 1)
           sht.Range(del_name).Delete Shift:=xlShiftUp
           del_busho = all_address.Item(tex)(1, 2)
           sht.Range(del_busho).Delete Shift:=xlShiftUp
           del_address = all_address.Item(tex)(1, 3)
           sht.Range(del_address).Delete Shift:=xlShiftUp
'==================================削除==============================================
'==================================アドレス取得======================================
           bush_ad = busho_address(i, 2)
           last_c = sht.Range(bush_ad).Column
           last_cell = sht.Cells(Rows.Count, last_c).End(xlUp).Row + 1
           r_column = sht.Range(bush_ad).Column
'==================================アドレス取得======================================
'==================================追加==============================================
           sht.Cells(last_cell, r_column) = busho
           sht.Cells(last_cell, r_column + 1) = tex
           sht.Cells(last_cell, r_column + 2) = name_busho_meado.Item(tex)(1, 2)
'==================================追加==============================================
'==================================削除==============================================
           name_busho_meado.Item(tex)(1, 1) = busho
           all_address.Item(tex)(1, 2) = del_busho
'==================================削除==============================================
            Exit For
        End If
    Next i
    
    With ListBox2
        .AddItem tex
    End With
    
    Set sht = Nothing
    Set name_busho_meado = Nothing
    Set all_address = Nothing
    CommandButton1.Enabled = True  'ボタン連打防止

End Sub
Private Sub CommandButton2_Click()  '削除

    CommandButton2.Enabled = False 'ボタン連打防止
    Dim tex As String
    Dim sht As Worksheet  'マスターシートを参照する。
    Set sht = a.Mysht  'マスターシートを参照する。
    Dim list_name() As Variant
    Dim busho_ad As String
    Dim busho As String
    Dim name As String
    Dim meado As String
    Dim rc As Integer
    Dim curKey As Variant
    Set all_address = a.all_addresses
    tex = ListBox2.Text
    
    If tex = "" Then
          MsgBox "名簿から名前を選んで下さい。"
          CommandButton2.Enabled = True
          Exit Sub
    Else
          busho = all_address.Item(tex)(1, 2)
          name = all_address.Item(tex)(1, 1)
          meado = all_address.Item(tex)(1, 3)
    End If
    

    rc = MsgBox("本当に" & sht.Range(busho) & " " & sht.Range(name) & "を削除してもよろしいでしょうか。", vbYesNo + vbQuestion, "確認")
    If rc = vbNo Then
          CommandButton2.Enabled = True  'ボタン連打防止
          Exit Sub
    ElseIf rc = vbYes Then
          sht.Range(busho).Delete Shift:=xlShiftUp
          sht.Range(name).Delete Shift:=xlShiftUp
          sht.Range(meado).Delete Shift:=xlShiftUp
          all_address.Remove (tex)
    End If

    With ComboBox1
        busho = .List(.ListIndex)
    End With

    With ListBox2
        .Clear
          For Each curKey In all_address
                list_name = all_address.Item(curKey)
                busho_ad = list_name(1, 2)
                If busho = sht.Range(busho_ad) Then
                    .AddItem curKey
                End If
          Next
    End With
    
    Set sht = Nothing
    Set all_address = Nothing
    CommandButton2.Enabled = True 'ボタン連打防止

End Sub

Private Sub CommandButton3_Click()

 Unload Me
 
End Sub

Private Sub CommandButton4_Click()
        
 Unload Me
 組合員登録.Show

End Sub
Private Sub CommandButton5_Click() '編集

    CommandButton5.Enabled = False 'ボタン連打防止
    Dim tex As String
    Dim sht As Worksheet  'マスターシートを参照する。
    Set sht = a.Mysht  'マスターシートを参照する。
    Dim list_name() As Variant
    Dim busho_ad As String
    Dim busho_address As String
    Dim name_address As String
    Dim mead_address As String
    Dim rc As Integer
    Dim curKey As Variant
    Set all_address = a.all_addresses
    tex = ListBox2.Text
    
    If tex = "" Then
          MsgBox "名簿から名前を選んで下さい。"
          CommandButton5.Enabled = True
          Exit Sub
    Else
          busho_address = all_address.Item(tex)(1, 2)
          name_address = all_address.Item(tex)(1, 1)
          mead_address = all_address.Item(tex)(1, 3)
    End If
    
          busho = sht.Range(busho_address)
          name_edit = sht.Range(name_address)
          mead_edit = sht.Range(mead_address)
          編集中.Show 'ユーザー情報受け取り
          sht.Range(busho_address) = busho
          sht.Range(name_address) = name_edit
          sht.Range(mead_address) = mead_edit
          Set all_address = a.all_addresses
          
    With ComboBox1
        busho = .List(.ListIndex)
    End With

    With ListBox2
        .Clear
          For Each curKey In all_address
                list_name = all_address.Item(curKey)
                busho_ad = list_name(1, 2)
                If busho = sht.Range(busho_ad) Then
                    .AddItem curKey
                End If
          Next
    End With
    
    Set sht = Nothing
    Set all_address = Nothing
    CommandButton5.Enabled = True 'ボタン連打防止

End Sub

Private Sub CommandButton6_Click()

    Dim i As Long
    Dim rc As Integer
    
    If TextBox1.Text = "" Then Exit Sub
    For i = 0 To ListBox3.ListCount - 1             ''(1)
        If ListBox3.List(i) = TextBox1.Text Then
              rc = MsgBox(ListBox3.List(i) & "が見つかりました。検索を続けますか。", vbYesNo + vbQuestion, "確認")
              ListBox3.Selected(i) = True
              If rc = vbNo Then Exit Sub
        ElseIf Left(ListBox3.List(i), 2) = Left(TextBox1.Text, 2) Then
              rc = MsgBox(ListBox3.List(i) & "が見つかりました。検索を続けますか。", vbYesNo + vbQuestion, "確認")
              ListBox3.Selected(i) = True
              If rc = vbNo Then Exit Sub
        ElseIf Left(ListBox3.List(i), 3) = Left(TextBox1.Text, 3) Then  ''(2)
              rc = MsgBox(ListBox3.List(i) & "が見つかりました。検索を続けますか。", vbYesNo + vbQuestion, "確認")
              ListBox3.Selected(i) = True
              If rc = vbNo Then Exit Sub
        End If
    Next i
    MsgBox "検索が終了しました。"

End Sub

Private Sub UserForm_Terminate()
    Set a = Nothing
    Application.ScreenUpdating = True
End Sub
Private Sub UserForm_initialize()
    Application.ScreenUpdating = False
    Call GETADD_DELETE
    Dim i As Integer
    Dim j As Integer
    Dim curKey As Variant
    Dim list_name() As Variant
    Dim name_ad As String
    Dim sht As Worksheet
    Set sht = a.Mysht
    Dim busho_ad As String
    Set all_address = a.all_addresses
    Set busho_all_address = a.busho_all_address
    list_address = busho_all_address.items
    list_busho = busho_all_address.keys
    With ComboBox1 'リストボックスの表示'
        For i = 0 To UBound(list_busho)
            .AddItem list_busho(i)
        Next i
    End With

    With ListBox3
         .Clear
          For Each curKey In all_address
                list_name = all_address.Item(curKey)
                name_ad = list_name(1, 1)
                .AddItem sht.Range(name_ad)
          Next
    End With
    
    Label2.Caption = Person
    
End Sub
