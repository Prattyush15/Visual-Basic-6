VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12195
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   12195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcheck 
      Caption         =   "Check"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   600
      TabIndex        =   4
      Top             =   4680
      Width           =   2235
   End
   Begin VB.TextBox txtguess 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   540
      TabIndex        =   3
      Top             =   2820
      Width           =   2775
   End
   Begin VB.CommandButton cmdrnum 
      Caption         =   "Random Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label 
      Caption         =   "Random number selected"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   1
      Left            =   3420
      TabIndex        =   5
      Top             =   780
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Take a Guess"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   3435
   End
   Begin VB.Label lblrnum 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7140
      TabIndex        =   1
      Top             =   780
      Visible         =   0   'False
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rnum As Integer
Sub newgame()
txtguess = ""
lblrnum = ""
Label(1).Visible = False
End Sub
Private Sub cmdcheck_Click()
Dim i As Integer
Dim g1 As Integer
i = False
g1 = Int(txtguess)
If lblrnum = "" Then
    i = MsgBox("Please generate a number before you start checking")
    Exit Sub
End If
Do While i = False
    If g1 = lblrnum Then
        i = MsgBox("You got it right. The correct number is " + lblrnum)
        i = MsgBox("Would you like to start a new game?", vbYesNo)
        If i = vbNo Then
            i = MsgBox("Thank you for playing")
            End
        ElseIf i = vbYes Then
            newgame
        End If
    ElseIf g1 > lblrnum Then
        i = MsgBox("The number is lower than " + txtguess)
    ElseIf g1 < lblrnum Then
        i = MsgBox("The number is higher than " + txtguess)
    End If
Loop
txtguess = ""
End Sub

Private Sub cmdrnum_Click()
Randomize
Label(1).Visible = True
rnum = Int(Rnd * 100) + 1
lblrnum = rnum
End Sub

Private Sub lblhi_Click()
End Sub

Private Sub lblrnum_Click()

End Sub
