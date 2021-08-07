VERSION 5.00
Begin VB.Form Multiples 
   Caption         =   "Form1"
   ClientHeight    =   8250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11565
   LinkTopic       =   "Form1"
   ScaleHeight     =   8250
   ScaleWidth      =   11565
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstcount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      ItemData        =   "Multiples.frx":0000
      Left            =   180
      List            =   "Multiples.frx":0002
      TabIndex        =   10
      Top             =   3060
      Width           =   3975
   End
   Begin VB.CommandButton cmdclear 
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
      Height          =   1095
      Left            =   10080
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox multiline 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   4500
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Top             =   4560
      Width           =   5355
   End
   Begin VB.TextBox txtmult 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3120
      TabIndex        =   6
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdend 
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
      Height          =   1335
      Left            =   10080
      TabIndex        =   3
      Top             =   540
      Width           =   1335
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   7860
      TabIndex        =   2
      Top             =   540
      Width           =   2115
   End
   Begin VB.TextBox txtnum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   60
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label 
      Caption         =   "Number of Multiples"
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
      Index           =   2
      Left            =   3120
      TabIndex        =   7
      Top             =   300
      Width           =   3675
   End
   Begin VB.Label lblsum 
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
      Height          =   1215
      Left            =   6540
      TabIndex        =   5
      Top             =   2700
      Width           =   2175
   End
   Begin VB.Label Label 
      Caption         =   "Sum:"
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
      Index           =   1
      Left            =   4860
      TabIndex        =   4
      Top             =   2940
      Width           =   1575
   End
   Begin VB.Label Label 
      Caption         =   "Enter Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2235
   End
End
Attribute VB_Name = "Multiples"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdclear_Click()
txtnum = ""
txtmult = ""
lstcount.Clear
multiline = ""
lblsum = ""
txtnum.Enabled = True
txtmult.Enabled = True
lstcount.Enabled = True
multiline.Enabled = True
lblsum.Enabled = True
txtnum.SetFocus
End Sub
Private Sub cmdend_Click()
End
End Sub
Private Sub cmdfind_Click()
Dim num As Long
Dim count As Long
Dim Sum As Long
Dim mult As Long
Dim strcount As String
mult = Val(txtmult)
strcount = ""
Sum = 0
num = Val(txtnum)
Dim number As Long

lstcount.AddItem num
For count = 1 To mult - 1
    num = num + lstcount.List(0)
    
    lstcount.AddItem num
    
    strcount = strcount + Str(num) + " "
    Sum = Sum + num
Next count
    multiline = Str(txtnum) + " " + strcount
    lblsum = Sum + txtnum
cmdclear.SetFocus

End Sub

Private Sub txtmult_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdfind.SetFocus
End If
End Sub

Private Sub txtnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtmult.SetFocus
End If
End Sub


