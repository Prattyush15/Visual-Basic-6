VERSION 5.00
Begin VB.Form CerealProject 
   BackColor       =   &H80000007&
   Caption         =   "Cereal"
   ClientHeight    =   6060
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3420
      Top             =   4680
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   3480
      Top             =   5280
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      TabIndex        =   6
      Top             =   5100
      Width           =   2955
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      TabIndex        =   5
      Top             =   3540
      Width           =   2955
   End
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check Supply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   4800
      TabIndex        =   4
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtbowls 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4800
      TabIndex        =   3
      Top             =   1140
      Width           =   2895
   End
   Begin VB.TextBox txtboxes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label lblBuy 
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Buy more cereal!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "I need some ASAP"
      Top             =   3060
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label lblOK 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Cereal supply is OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Don't buy anymore"
      Top             =   4320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label 
      Caption         =   "Number of bowls eaten per week."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label 
      Caption         =   "Number of boxes on hand. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   180
      Width           =   2955
   End
End
Attribute VB_Name = "CerealProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheck_Click()
 Dim Boxes As Integer, Bowls As Integer
    Dim Servings As Integer
    Boxes = Val(txtboxes)
    Bowls = Val(txtbowls)
    Servings = Boxes * 12
    If Servings >= Bowls * 2 Then
        lblOK.Visible = True
    Else
        lblBuy.Visible = True
    End If
    cmdClear.SetFocus
    
End Sub

Private Sub cmdClear_Click()
txtboxes = ""
txtbowls = ""
lblBuy.Visible = False
lblOK.Visible = False
txtboxes.SetFocus
End Sub

Private Sub cmdQuit_Click()
End
End Sub

Private Sub Timer_Timer()
Dim nHeight As Integer
Dim n As Integer
n = Second(Now) Mod 10
If n = 0 Then
    nHeight = 10
ElseIf n = 1 Or n = 9 Then
    nHeight = 25
ElseIf n = 2 Or n = 8 Then
    nHeight = 10
ElseIf n = 3 Or n = 7 Then
    nHeight = 25
ElseIf n = 4 Or n = 6 Then
    nHeight = 10
ElseIf n = 5 Then
    nHeight = 25
End If
lblBuy.FontSize = nHeight
End Sub

Private Sub Timer1_Timer()
Dim nHeight As Integer
Dim n As Integer
n = Second(Now) Mod 10
If n = 0 Then
    nHeight = 10
ElseIf n = 1 Or n = 9 Then
    nHeight = 20
ElseIf n = 2 Or n = 8 Then
    nHeight = 10
ElseIf n = 3 Or n = 7 Then
    nHeight = 20
ElseIf n = 4 Or n = 6 Then
    nHeight = 10
ElseIf n = 5 Then
    nHeight = 20
End If
lblOK.FontSize = nHeight
End Sub

Private Sub txtbowls_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdCheck.SetFocus
End If
End Sub

Private Sub txtboxes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtbowls.SetFocus
End If
End Sub
