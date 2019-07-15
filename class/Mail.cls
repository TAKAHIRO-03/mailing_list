VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit '絶対変数宣言してね！って意味
'======================================================================
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'↑使用するパソコンによって、変更しなきゃいけない可能性有り
'======================================================================
Private wkNSes As Object
Private wkNDB As Object
Private wkNDoc As Object
Private wkNRtItem As Object
Private wkNAtt As Object
Private AttFName As String
Private busho() As String
Private names() As String
Private adress() As String
Private lists() As String
Private curKey As Variant
Private i As Long
Private j As Long
Private k As Long

' 初期化処理
Private Sub Class_Initialize()

    Set wkNSes = CreateObject("Notes.NotesSession")
    Set wkNDB = wkNSes.GETDATABASE("", "")
    wkNDB.OpenMail
    Debug.Print "mail初期化したよ"
    
End Sub
' 終了時処理
Private Sub Class_Terminate()

    Set wkNAtt = Nothing
    Set wkNRtItem = Nothing
    Set wkNDoc = Nothing
    Set wkNDB = Nothing
    Set wkNSes = Nothing
    Debug.Print "mail終わったよ"
    
End Sub

Public Function submit_mail(name_busho_meado As Object, Contents() As String, Titles As String)
       
    ReDim busho(name_busho_meado.Count - 1)
    ReDim names(name_busho_meado.Count - 1)
    ReDim adress(name_busho_meado.Count - 1)
    Dim Title As String
    Dim Content() As String
    Dim r As Integer
    Dim Quest As String
    Dim log_date As String
    flag = 1
    
    Title = Titles
    Content = Contents
    i = 0
    For Each curKey In name_busho_meado
         lists = name_busho_meado.Item(curKey)
         names(i) = curKey
         busho(i) = lists(1, 1)
         adress(i) = lists(1, 2)
         i = i + 1
    Next
    
    Quest = InputBox("アンケート内容を入力して下さい。" & vbLf & "（例：三木会、統一労懇）", _
                       "アンケート内容")
        
    'キャンセルされたかどうかの判断はStrPtr関数で行います。
    If StrPtr(Quest) = 0 Then
        'キャンセルボタンか×ボタンが押下された場合
        MsgBox "入力がキャンセルされました。", vbExclamation
    ElseIf Quest = "" Then
        '値を入力しないでOKボタンを押下した場合
        MsgBox "値が未入力です。", vbExclamation
    End If
    
    For j = 0 To UBound(adress)
        Set wkNDoc = wkNDB.CREATEDOCUMENT()
        wkNDoc.Subject = Title
        wkNDoc.SendTo = Array(adress(j))
        'wkNDoc.CopyTo = Array("xxx@xxx")
        'wkNDoc.blindCopyTo = Array("xxx@xxx")
        Set wkNRtItem = wkNDoc.CreateRichTextItem("BODY")
        With wkNRtItem
                .APPENDTEXT busho(j) '部署
                .ADDNEWLINE 2
                .APPENDTEXT names(j) & "　様" '宛名
                .ADDNEWLINE 2
            For k = 0 To UBound(Content)
                .APPENDTEXT Content(k)
                .ADDNEWLINE 1
            Next k
        End With
        Sleep 500
                
        'OSに処理を渡す。
        DoEvents
        処理中.Label1.Caption = "処理中です..." & j + 1 & "件"
        
        wkNDoc.Send False
        Set wkNRtItem = Nothing
        Set wkNDoc = Nothing
        
        r = ThisWorkbook.Worksheets("Log").Cells(Rows.Count, 1).End(xlUp).Row + 1
        log_date = Now
        
        With ThisWorkbook.Worksheets("Log")
            .Cells(r, 1) = j + 1
            .Cells(r, 2) = Quest
            .Cells(r, 3) = busho(j)
            .Cells(r, 4) = names(j)
            .Cells(r, 5) = adress(j)
            .Cells(r, 6) = log_date
            .Cells(r, 7) = Person
        End With
        
        If flag = 0 Then '処理中の閉じる検知
            wkNDoc.Send False
            Set wkNRtItem = Nothing
            Set wkNDoc = Nothing
            Exit For
        End If
    
    Next j
    
        MsgBox "送信" & j & "件" & "完了しました。", vbOKOnly + vbInformation

End Function
