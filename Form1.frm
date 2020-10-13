VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memeriksa Cancel pada InputBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5355
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5355
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coding ini menggunakan fungsi StrPtr (string pointer), 'yang mengembalikan pointer ke string.
'Jika user menekan Cancel pada InputBox, hasil dari 'InputBox akan menjadi pointer yang menunjuk kepada "Null" (vbNullString Constant), dan itu sama dengan 0.
'Catatan: string kosong ("") BUKAN Null.

Private Sub Form_Load()
    Dim str As String
    str = InputBox("Press OK or Cancel")
    If StrPtr(str) = 0 Then
       MsgBox "Anda menekan Cancel", vbInformation, _
              "Cancel"
    End If
End Sub


