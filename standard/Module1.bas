Attribute VB_Name = "Module1"
Option Explicit
Public g As Getinfo
Public m As Mail
Public a As Add_Delete
Public Person As String
Public adress As String
Public busho As String
Public name_edit As String
Public mead_edit As String
Public busho_address() As Variant
Public list_busho() As Variant
Public list_address() As Variant
Public Content() As String
Public name_busho_meado As Object
Public all_address As Object
Public busho_all_address As Object
Public Department As String
Public Title As String
Sub GETINFO_CLASS()
  Set g = New Getinfo
End Sub
Sub GETMAIL_CLASS()
    Set m = New Mail
End Sub
Sub GETADD_DELETE()
    Set a = New Add_Delete
End Sub
