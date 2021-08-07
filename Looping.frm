VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   ScaleHeight     =   9135
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEnd 
      Caption         =   "End"
      Height          =   735
      Left            =   12360
      TabIndex        =   7
      Top             =   1020
      Width           =   795
   End
   Begin VB.TextBox txtmulti 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3555
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4560
      Width           =   6975
   End
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
      Height          =   4410
      ItemData        =   "Looping.frx":0000
      Left            =   180
      List            =   "Looping.frx":0002
      TabIndex        =   5
      Top             =   3300
      Width           =   3315
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
      Height          =   1155
      Left            =   8100
      TabIndex        =   2
      Top             =   780
      Width           =   3075
   End
   Begin VB.TextBox txtnum 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   60
      TabIndex        =   1
      Top             =   1500
      Width           =   3735
   End
   Begin VB.Label lblsum 
      Alignment       =   2  'Center
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
      Height          =   1155
      Left            =   7860
      TabIndex        =   4
      Top             =   3060
      Width           =   3555
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
      Height          =   915
      Index           =   1
      Left            =   6060
      TabIndex        =   3
      Top             =   3300
      Width           =   1515
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
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
      Width           =   3675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEnd_Click()
End
End Sub

Private Sub cmdfind_Click()
Dim num As Integer
Dim count As Integer
Dim sum As Long
Dim strcount As String
strcount = ""
num = Val(txtnum.Text)
For count = 1 To num
    lstcount.AddItem count
    strcount = strcount + Str(count) + " "
    sum = sum + count
Next count
txtmulti = strcount
lblsum = sum
End Sub
