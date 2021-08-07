VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000007&
   Caption         =   "Form1"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   ScaleHeight     =   7515
   ScaleWidth      =   12180
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      Height          =   255
      Left            =   300
      TabIndex        =   28
      Top             =   180
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer Timer 
      Interval        =   50
      Left            =   2040
      Top             =   6060
   End
   Begin VB.Label lblright 
      Height          =   7515
      Left            =   12000
      TabIndex        =   27
      Top             =   0
      Width           =   195
   End
   Begin VB.Label lblleft 
      Height          =   7515
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   135
   End
   Begin VB.Label lblwall 
      Height          =   75
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label y6 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9420
      TabIndex        =   24
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label y5 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7740
      TabIndex        =   23
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label y4 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6060
      TabIndex        =   22
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label y3 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4380
      TabIndex        =   21
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label y2 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2700
      TabIndex        =   20
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label y1 
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1020
      TabIndex        =   19
      Top             =   2940
      Width           =   1695
   End
   Begin VB.Label g6 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9420
      TabIndex        =   18
      Top             =   2580
      Width           =   1695
   End
   Begin VB.Label g5 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7740
      TabIndex        =   17
      Top             =   2580
      Width           =   1695
   End
   Begin VB.Label g4 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6060
      TabIndex        =   16
      Top             =   2580
      Width           =   1695
   End
   Begin VB.Label g3 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4380
      TabIndex        =   15
      Top             =   2580
      Width           =   1695
   End
   Begin VB.Label g2 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2700
      TabIndex        =   14
      Top             =   2580
      Width           =   1695
   End
   Begin VB.Label g1 
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1020
      TabIndex        =   13
      Top             =   2580
      Width           =   1695
   End
   Begin VB.Label b6 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9420
      TabIndex        =   12
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label b5 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7740
      TabIndex        =   11
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label b4 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6060
      TabIndex        =   10
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label b3 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4380
      TabIndex        =   9
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label b2 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2700
      TabIndex        =   8
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label b1 
      BackColor       =   &H00C0C000&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Index           =   0
      Left            =   1020
      TabIndex        =   7
      Top             =   2220
      Width           =   1695
   End
   Begin VB.Label r6 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   9420
      TabIndex        =   6
      Top             =   1860
      Width           =   1695
   End
   Begin VB.Label r5 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   7740
      TabIndex        =   5
      Top             =   1860
      Width           =   1695
   End
   Begin VB.Label r4 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   6060
      TabIndex        =   4
      Top             =   1860
      Width           =   1695
   End
   Begin VB.Label r3 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   4380
      TabIndex        =   3
      Top             =   1860
      Width           =   1695
   End
   Begin VB.Label r2 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   2700
      TabIndex        =   2
      Top             =   1860
      Width           =   1695
   End
   Begin VB.Label r1 
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1020
      TabIndex        =   1
      Top             =   1860
      Width           =   1695
   End
   Begin VB.Image imgball 
      Height          =   180
      Left            =   5640
      Picture         =   "brick.frx":0000
      Top             =   3660
      Width           =   180
   End
   Begin VB.Label lblpaddle 
      BackColor       =   &H8000000D&
      Height          =   375
      Left            =   4140
      TabIndex        =   0
      Top             =   7140
      Width           =   3795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bsx As Integer
Dim bsy As Integer

Private Sub Timer_Timer()
imgball.Move imgball.Left + bsx * 70, imgball.Top + bsy * 70
If imgball.Left + imgball.Width + 200 > Width - 100 Then
    bsx = bsx * -1
End If
If imgball.Top + imgball.Height < lblwall.Top Then
    bsy = bsy * -1
End If
If imgball.Top + imgball.Height > lblpaddle.Top + lblpaddle.Height Then
    cmdnew.Visible = True
End If
If imgball.Left > lblpaddle.Left Then
    If imgball.Left + imgball.Width < lblpaddle.Left + lblpaddle.Width Then
        If imgball.Top + imgball.Height > lblpaddle.Top Then
            bsy = bsy * -1
        End If
    End If
End If
If imgball.Top > lblpaddle.Top Then
    bsy = bsy * -1
End If
If imgball.Top < y1.Top + y1.Left Then
    y1.Visible = False
End If
If imgball.Top < y2.Top + y2.Left Then
    y1.Visible = False
End If
If imgball.Top < y3.Top + y3.Left Then
    y1.Visible = False
End If
If imgball.Top < y4.Top + y4.Left Then
    y1.Visible = False
End If
If imgball.Top < y5.Top + y5.Left Then
    y1.Visible = False
End If
If imgball.Top < y6.Top + y6.Left Then
    y1.Visible = False
End If

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblpaddle.Left = X
End Sub
Private Sub Form_Load()
bsx = 1
bsy = 1
End Sub
