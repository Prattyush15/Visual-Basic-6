VERSION 5.00
Begin VB.Form Calculator 
   BackColor       =   &H00C0C000&
   Caption         =   "Calculator"
   ClientHeight    =   8025
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8025
   ScaleWidth      =   13200
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture 
      Height          =   3195
      Index           =   1
      Left            =   9960
      Picture         =   "Calculator.frx":0000
      ScaleHeight     =   3135
      ScaleWidth      =   3195
      TabIndex        =   18
      Top             =   4860
      Width           =   3255
   End
   Begin VB.CommandButton cmdDivide 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4740
      TabIndex        =   5
      ToolTipText     =   "Divides number 1 by number 2"
      Top             =   2940
      Width           =   1035
   End
   Begin VB.PictureBox Picture 
      Height          =   3255
      Index           =   0
      Left            =   0
      Picture         =   "Calculator.frx":1A4A
      ScaleHeight     =   3195
      ScaleWidth      =   3195
      TabIndex        =   17
      Top             =   4800
      Width           =   3255
   End
   Begin VB.CommandButton cmdquit 
      BackColor       =   &H80000005&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4800
      TabIndex        =   11
      ToolTipText     =   "Quits the program"
      Top             =   6660
      Width           =   3555
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   4800
      MaskColor       =   &H000080FF&
      TabIndex        =   10
      ToolTipText     =   "Clears all numbers entered"
      Top             =   5340
      Width           =   3555
   End
   Begin VB.CommandButton cmdsquare2 
      BackColor       =   &H80000005&
      Caption         =   "^2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   11520
      TabIndex        =   9
      ToolTipText     =   "Squares the first number by two. The second number doesn't do anything."
      Top             =   3060
      Width           =   1275
   End
   Begin VB.CommandButton cmdpower10 
      Caption         =   "10^x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10140
      TabIndex        =   8
      ToolTipText     =   "First number to the power of ten. Uses only the first number, second number doesn't do anything"
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton cmdavg 
      Caption         =   "Avg"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8700
      TabIndex        =   7
      ToolTipText     =   "Averages the 2 numbers"
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton cmdsquare 
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   7500
      MaskColor       =   &H80000005&
      TabIndex        =   6
      ToolTipText     =   "Uses the first number and squares it to the second number"
      Top             =   3060
      Width           =   975
   End
   Begin VB.CommandButton cmdmultiply 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3420
      TabIndex        =   4
      ToolTipText     =   "Multiplies number 1 and 2"
      Top             =   2940
      Width           =   1095
   End
   Begin VB.CommandButton cmdsubtract 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   1380
      TabIndex        =   3
      ToolTipText     =   "Subtracts number 1 and 2"
      Top             =   2880
      Width           =   1035
   End
   Begin VB.CommandButton cmdadd 
      BackColor       =   &H80000005&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   180
      TabIndex        =   2
      ToolTipText     =   "Adds number 1 and 2"
      Top             =   2880
      Width           =   1035
   End
   Begin VB.TextBox txtnum2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3300
      TabIndex        =   1
      Top             =   1380
      Width           =   2655
   End
   Begin VB.TextBox txtnum1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   180
      TabIndex        =   0
      Top             =   1440
      Width           =   2475
   End
   Begin VB.Label lblanswer 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   6840
      TabIndex        =   16
      ToolTipText     =   "The Answer"
      Top             =   1440
      Width           =   6015
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Answer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   3
      Left            =   8820
      TabIndex        =   15
      Top             =   960
      Width           =   1710
   End
   Begin VB.Label Label 
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Number 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3300
      TabIndex        =   14
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label 
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Number 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   180
      TabIndex        =   13
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Calculator"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Index           =   0
      Left            =   5460
      TabIndex        =   12
      Top             =   0
      Width           =   2880
   End
End
Attribute VB_Name = "Calculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdsqrt_Click()

End Sub

Private Sub cmdadd_Click()
Dim Num1 As Double
Dim Num2 As Double
Dim Answer As Double
Num1 = Val(txtnum1.Text)
Num2 = Val(txtnum2.Text)
Answer = Num1 + Num2
lblanswer.Caption = Answer
cmdclear.SetFocus
End Sub

Private Sub cmdavg_Click()
Dim Num1 As Double
Dim Num2 As Double
Dim Answer As Double
Num1 = Val(txtnum1.Text)
Num2 = Val(txtnum2.Text)
Answer = (Num1 + Num2) / 2
lblanswer.Caption = Answer
cmdclear.SetFocus
End Sub

Private Sub cmdclear_Click()
txtnum1 = ""
txtnum2 = ""
lblanswer = ""
txtnum1.SetFocus

End Sub

Private Sub txtanswer_Change()

End Sub

Private Sub cmddivision_Click()
Dim Num1 As Double
Dim Num2 As Double
Dim Answer As Double
Num1 = Val(txtnum1.Text)
Num2 = Val(txtnum2.Text)
If Num2 > 0 Then
    Answer = (Num1 / Num2)
End If
If Num2 < 0 Then
    Answer = (Num1 / Num2)
End If
If Num2 = 0 Then
    Answer = "Undefined"
End If
lblanswer.Caption = Answer
cmdclear.SetFocus

End Sub

Private Sub cmdDivide_Click()
Dim Num1 As Double
Dim Num2 As Double
Dim Answer As String
Num1 = Val(txtnum1.Text)
Num2 = Val(txtnum2.Text)
If Num2 > 0 Then
    Answer = Num1 / Num2
End If
If Num2 < 0 Then
    Answer = Num1 / Num2
End If
If Num2 = 0 Then
    Answer = "Undefined"
End If
lblanswer.Caption = Answer
cmdclear.SetFocus

End Sub

Private Sub cmdmultiply_Click()
Dim Num1 As Double
Dim Num2 As Double
Dim Answer As Double
Num1 = Val(txtnum1.Text)
Num2 = Val(txtnum2.Text)
Answer = Num1 * Num2
lblanswer.Caption = Answer
cmdclear.SetFocus
End Sub

Private Sub cmdpower10_Click()
Dim Num1 As Double
Dim Num2 As Double
Dim Answer As String
Num1 = Val(txtnum1.Text)
Num2 = Val(txtnum2.Text)
Answer = 10 ^ Num1
lblanswer.Caption = Answer
cmdclear.SetFocus
End Sub

Private Sub cmdquit_Click()
End
End Sub

Private Sub cmdsquare_Click()
Dim Num1 As Single
Dim Num2 As Single
Dim Answer As Double
Num1 = Val(txtnum1.Text)
Num2 = Val(txtnum2.Text)
Answer = Num1 ^ Num2
lblanswer.Caption = Answer
cmdclear.SetFocus
End Sub

Private Sub cmdsquare2_Click()
Dim Num1 As Double
Dim Num2 As Double
Dim Answer As Double
Num1 = Val(txtnum1.Text)
Num2 = Val(txtnum2.Text)
Answer = Num1 ^ 2
lblanswer.Caption = Answer
cmdclear.SetFocus
End Sub

Private Sub cmdsubtract_Click()
Dim Num1 As Double
Dim Num2 As Double
Dim Answer As Double
Num1 = Val(txtnum1.Text)
Num2 = Val(txtnum2.Text)
Answer = Num1 - Num2
lblanswer.Caption = Answer
cmdclear.SetFocus
End Sub

Private Sub pic_Click()

End Sub

