VERSION 5.00
Begin VB.Form JankAlpha 
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   3
      Left            =   2940
      TabIndex        =   4
      Top             =   2100
      Width           =   5115
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   2
      Left            =   2940
      TabIndex        =   3
      Top             =   1440
      Width           =   5115
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   2940
      TabIndex        =   2
      Top             =   780
      Width           =   5115
   End
   Begin VB.TextBox Text 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   2940
      TabIndex        =   1
      Top             =   120
      Width           =   5115
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "End"
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
      Left            =   8340
      TabIndex        =   7
      Top             =   1800
      Width           =   1335
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
      Height          =   795
      Left            =   8340
      TabIndex        =   6
      Top             =   1020
      Width           =   1335
   End
   Begin VB.CommandButton cmdfind 
      Caption         =   "Find"
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
      Left            =   8340
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label 
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
      Height          =   975
      Index           =   3
      Left            =   7500
      TabIndex        =   15
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label Label 
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
      Height          =   975
      Index           =   2
      Left            =   5040
      TabIndex        =   14
      Top             =   3360
      Width           =   2235
   End
   Begin VB.Label Label 
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
      Height          =   975
      Index           =   1
      Left            =   2520
      TabIndex        =   13
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label 
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
      Height          =   975
      Index           =   0
      Left            =   60
      TabIndex        =   12
      Top             =   3360
      Width           =   2115
   End
   Begin VB.Label faksjdhlfakjsdh 
      Alignment       =   2  'Center
      Caption         =   "Alphabetical Order"
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
      Left            =   60
      TabIndex        =   11
      Top             =   2700
      Width           =   9735
   End
   Begin VB.Label dsfgsdfg 
      Caption         =   "Enter Word 4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   4
      Left            =   60
      TabIndex        =   10
      Top             =   2160
      Width           =   2475
   End
   Begin VB.Label dsfgdfsg 
      Caption         =   "Enter Word 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   3
      Left            =   60
      TabIndex        =   9
      Top             =   1500
      Width           =   2415
   End
   Begin VB.Label dsfgsdfg 
      Caption         =   "Enter Word 2"
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
      Index           =   1
      Left            =   60
      TabIndex        =   8
      Top             =   840
      Width           =   2955
   End
   Begin VB.Label fdgdfgsdf 
      Caption         =   "Enter Word 1"
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
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   2955
   End
End
Attribute VB_Name = "JankAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclear_Click()
Text(0) = ""
Text(1) = ""
Text(2) = ""
Text(3) = ""
Label(0) = ""
Label(1) = ""
Label(2) = ""
Label(3) = ""
Text(0).SetFocus
End Sub

Private Sub cmdend_Click()
End
End Sub

Private Sub cmdfind_Click()
Dim w1 As String
Dim w2 As String
Dim w3 As String
Dim w4 As String
w1 = Text(0).Text
w2 = Text(1).Text
w3 = Text(2).Text
w4 = Text(3).Text
Sort
Text(0).Text = w1
Text(1).Text = w2
Text(2).Text = w3
Text(3).Text = w4
cmdclear.SetFocus
End Sub
Private Sub Cmdquit_Click()
End

End Sub

Sub Sort()

Dim tempInt As Integer
Dim i As Integer
Dim j As Integer

Trim (Label(0))
Trim (Label(1))
Trim (Label(2))
Trim (Label(3))
For i = 0 To 3
    For j = 0 To 3
Trim (Text(i))
  tempInt = StrComp(UCase(Text(i).Text), UCase(Text(j).Text))
    
    If (tempInt = -1) Then
        Dim temprStr As String
        
       tempStr = Text(i).Text
      
     Text(i).Text = Text(j).Text
    
    Text(j).Text = tempStr
  End If
    
  Next j
Next i

For i = 0 To 3
   Label(i).Caption = Text(i).Text

Next i

End Sub


