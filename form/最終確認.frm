VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 最終確認 
   Caption         =   "最終確認"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7575
   OleObjectBlob   =   "最終確認.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "最終確認"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    Dim rc As Integer
    rc = MsgBox("本当にメールを送信してもよろしいでしょうか。", vbYesNo + vbQuestion, "確認")
    
    If rc = vbNo Then
          Exit Sub
    ElseIf rc = vbYes Then
         処理中.Show
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
    Label11.Caption = name_busho_meado.Count & "名"

    With ListBox4
            .AddItem ""
            .AddItem "部署　※自動で入力されます。"
            .AddItem ""
            .AddItem "氏名　※自動で入力されます。"
            .AddItem ""
        For j = 0 To UBound(Content)
            .AddItem Content(j)
        Next j
    End With
    
End Sub




