VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Getinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '絶対変数宣言してね！って意味
Private name_busho_meado As Object
Private all_address As Object
Private busho_address_hash As Object
Private spath As String
Private ex      As New Excel.Application    '// 処理用Excel
Private wb      As Workbook                 '// ワークブック
Private sht     As Worksheet                '// 参照シート
Private cel     As Range                    '// 参照セル
Private i       As Integer
Private j       As Integer
Private last_cell As Long
Private migihaji As Integer
Private busho_name() As String
Private all_address_list() As String
Private busho_list() As Variant
Private busho As String
Private kii As String
Private busho_count As Integer
' 初期化処理
Private Sub Class_Initialize()

    Set sht = ThisWorkbook.Worksheets("名簿") '// 一番左のシートの入力セル範囲を取得
    Set cel = sht.Cells(1, 1)
    migihaji = sht.Cells(1, Columns.Count).End(xlToLeft).Column
    Debug.Print "Getinfo初期化したよ"
    
End Sub
' 終了時処理
Private Sub Class_Terminate()

    Set sht = Nothing
    Set cel = Nothing
    Set name_busho_meado = Nothing
    Set all_address = Nothing

    Debug.Print "Getinfo終わったよ"
    
End Sub
Public Function name_busho_meados() As Object
    Set name_busho_meados = CreateObject("Scripting.Dictionary")
    Set name_busho_meado = CreateObject("Scripting.Dictionary") 'ハッシ生成
    
    For i = 1 To migihaji Step 3
         Set cel = sht.Cells(1, i)
         last_cell = sht.Cells(Rows.Count, i).End(xlUp).Row
             For j = 1 To last_cell
               ReDim busho_name(1 To 1, 1 To 2)
               Set cel = sht.Cells(j, 1 + i)
               kii = cel
               Set cel = sht.Cells(j, i)
               busho_name(1, 1) = cel
               Set cel = sht.Cells(j, 2 + i)
               busho_name(1, 2) = cel
               name_busho_meado.Add kii, busho_name
             Next j
    Next i
    Set name_busho_meados = name_busho_meado
End Function
Public Function busho_lists() As Variant()

    ReDim busho_list(0)
    Set cel = sht.Cells(1, 1)
    Set busho_address_hash = CreateObject("Scripting.Dictionary")
    busho_list(0) = cel
    busho_count = 1
    
    For i = 1 To migihaji Step 3
        Set cel = sht.Cells(1, i)
        last_cell = sht.Cells(Rows.Count, i).End(xlUp).Row
            For j = 1 To last_cell
                Set cel = sht.Cells(j, i)
                If Not busho_list(busho_count - 1) = cel Then
                    ReDim Preserve busho_list(busho_count)
                    busho_list(busho_count) = cel
                    busho_count = busho_count + 1
                End If
            Next j
    Next i
'===============================重複削除用==========================================
    For k = 0 To UBound(busho_list)
        If busho_address_hash.Exists(busho_list(k)) = False Then
            busho_address_hash.Add busho_list(k), busho_list(k)
        End If
    Next k
'===============================重複削除用==========================================
    busho_lists = busho_address_hash.items
    Set busho_address_hash = Nothing
    Erase busho_list

End Function
Public Function all_addresses() As Object
    Set all_addresses = CreateObject("Scripting.Dictionary")
    Set all_address = CreateObject("Scripting.Dictionary") 'ハッシ生成
    
        For i = 1 To migihaji Step 3
         Set cel = sht.Cells(1, i)
         last_cell = sht.Cells(Rows.Count, i).End(xlUp).Row
             For j = 1 To last_cell
               ReDim all_address_list(1 To 1, 1 To 3)
               Set cel = sht.Cells(j, 1 + i)
               kii = cel
               all_address_list(1, 1) = cel.Address(False, False)
               Set cel = sht.Cells(j, i)
               all_address_list(1, 2) = cel.Address(False, False)
               Set cel = sht.Cells(j, 2 + i)
               all_address_list(1, 3) = cel.Address(False, False)
               all_address.Add kii, all_address_list
             Next j
    Next i

    Set all_addresses = all_address
    
End Function
