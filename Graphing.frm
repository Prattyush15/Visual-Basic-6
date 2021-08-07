VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic1 
      Height          =   8715
      Left            =   60
      ScaleHeight     =   8655
      ScaleWidth      =   13275
      TabIndex        =   0
      Top             =   0
      Width           =   13335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
pic1.Scale (-10, 10)-(10, -10)
pic1.Line (-10, 0)-(10, 0), RGB(255, 0, 0)
pic1.Line (0, -10)-(0, 10), RGB(0, 0, 255)
End Sub

Private Sub pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub
