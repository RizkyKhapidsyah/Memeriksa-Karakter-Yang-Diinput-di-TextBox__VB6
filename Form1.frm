VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memeriksa Karakter yang Diinput di TextBox"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1680
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text1_KeyPress(KeyAscii As Integer)
'1 = True
'0 = False
   MsgBox "Informasi karakter: " & vbCrLf & "Upper Case: " & IsCharUpper(KeyAscii) & vbCrLf & "Lower Case: " & IsCharLower(KeyAscii) & vbCrLf & "Alpha: " & IsCharAlpha(KeyAscii) & vbCrLf & "Alpha atau Numeric: " & IsCharAlphaNumeric(KeyAscii), vbInformation
End Sub


