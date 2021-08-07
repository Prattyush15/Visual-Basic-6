VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000D&
   Caption         =   "Form1"
   ClientHeight    =   10200
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   10200
   ScaleWidth      =   13500
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000B&
      ForeColor       =   &H8000000B&
      Height          =   10755
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   13515
      Begin VB.ListBox lstleader 
         Height          =   2010
         ItemData        =   "frogger.frx":0000
         Left            =   5160
         List            =   "frogger.frx":0002
         TabIndex        =   39
         Top             =   7980
         Visible         =   0   'False
         Width           =   5535
      End
      Begin VB.CommandButton cmdleaderboard 
         Caption         =   "Click here to see other peoples scores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   5100
         TabIndex        =   38
         Top             =   6540
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.CommandButton cmdenter 
         Caption         =   "Click here to start the game"
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
         Left            =   5100
         TabIndex        =   37
         Top             =   5340
         Width           =   5655
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000B&
         Caption         =   "Hint: Diagonal movement requires you to click 2 keys at the same time."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Index           =   5
         Left            =   0
         TabIndex        =   36
         Top             =   7680
         Width           =   2595
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000B&
         Caption         =   $"frogger.frx":0004
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4875
         Index           =   4
         Left            =   0
         TabIndex        =   35
         Top             =   2520
         Width           =   3135
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000B&
         Caption         =   "Instructions:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   60
         TabIndex        =   34
         Top             =   2100
         Width           =   1995
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Made By: Prattyush Giriraj"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   2
         Left            =   0
         TabIndex        =   33
         Top             =   1380
         Width           =   13515
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "A Very Difficult Version of Frogger"
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
         Index           =   1
         Left            =   0
         TabIndex        =   32
         Top             =   780
         Width           =   13515
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Super Frogger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   30
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   795
         Left            =   -60
         TabIndex        =   31
         Top             =   0
         Width           =   13635
      End
   End
   Begin VB.Frame Framefrog 
      Height          =   600
      Left            =   6600
      TabIndex        =   6
      Top             =   9600
      Width           =   600
      Begin VB.PictureBox frog 
         Height          =   600
         Left            =   0
         Picture         =   "frogger.frx":0118
         ScaleHeight     =   540
         ScaleWidth      =   540
         TabIndex        =   7
         Top             =   0
         Width           =   600
      End
   End
   Begin VB.Timer cartimer 
      Interval        =   100
      Left            =   1140
      Top             =   9480
   End
   Begin VB.Timer frogtimer 
      Interval        =   100
      Left            =   480
      Top             =   9480
   End
   Begin VB.Label Label22 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   2400
      TabIndex        =   41
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label Label21 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   3000
      TabIndex        =   40
      Top             =   2400
      Width           =   600
   End
   Begin VB.Image car11 
      Height          =   435
      Left            =   2100
      Picture         =   "frogger.frx":0427
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image car10 
      Height          =   435
      Left            =   6120
      Picture         =   "frogger.frx":0621
      Top             =   6600
      Width           =   720
   End
   Begin VB.Image car9 
      Height          =   435
      Left            =   5280
      Picture         =   "frogger.frx":081B
      Top             =   8400
      Width           =   720
   End
   Begin VB.Image car8 
      Height          =   435
      Left            =   12180
      Picture         =   "frogger.frx":0A15
      Top             =   9000
      Width           =   720
   End
   Begin VB.Image car5 
      Height          =   435
      Left            =   9840
      Picture         =   "frogger.frx":0C0F
      Top             =   6600
      Width           =   720
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "<----1 Point --->"
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
      Left            =   9180
      TabIndex        =   29
      Top             =   1200
      Width           =   3135
   End
   Begin VB.Label lbl 
      BackStyle       =   0  'Transparent
      Caption         =   "<----2 Points"
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
      Left            =   2820
      TabIndex        =   28
      Top             =   600
      Width           =   2295
   End
   Begin VB.Image car7 
      Height          =   435
      Left            =   660
      Picture         =   "frogger.frx":0E09
      Top             =   7800
      Width           =   720
   End
   Begin VB.Image car6 
      Height          =   435
      Left            =   11400
      Picture         =   "frogger.frx":1003
      Top             =   7800
      Width           =   720
   End
   Begin VB.Image car3 
      Height          =   435
      Left            =   3480
      Picture         =   "frogger.frx":11FD
      Top             =   9000
      Width           =   720
   End
   Begin VB.Image car2 
      Height          =   435
      Left            =   9360
      Picture         =   "frogger.frx":13F7
      Top             =   8400
      Width           =   720
   End
   Begin VB.Image car4 
      Height          =   435
      Left            =   7740
      Picture         =   "frogger.frx":15F1
      Top             =   9000
      Width           =   720
   End
   Begin VB.Label Label19 
      BackColor       =   &H00800080&
      Height          =   600
      Left            =   0
      TabIndex        =   27
      Top             =   9600
      Width           =   13515
   End
   Begin VB.Label Label18 
      BackColor       =   &H00808080&
      Height          =   600
      Left            =   0
      TabIndex        =   26
      Top             =   6000
      Width           =   13515
   End
   Begin VB.Label Label17 
      BackColor       =   &H00404080&
      Height          =   600
      Left            =   0
      TabIndex        =   25
      Top             =   3000
      Width           =   13515
   End
   Begin VB.Label Label1 
      Caption         =   "Score:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   1035
   End
   Begin VB.Label Score 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   435
      Left            =   1020
      TabIndex        =   23
      Top             =   0
      Width           =   1635
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000004&
      Height          =   600
      Left            =   1200
      TabIndex        =   22
      Top             =   600
      Width           =   600
   End
   Begin VB.Label Label15 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   1200
      TabIndex        =   21
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label Label13 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   1800
      TabIndex        =   20
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label Label12 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   12600
      TabIndex        =   19
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label Label11 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   12600
      TabIndex        =   18
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label Label10 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   12000
      TabIndex        =   17
      Top             =   2400
      Width           =   600
   End
   Begin VB.Label Label9 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   7800
      TabIndex        =   16
      Top             =   1200
      Width           =   600
   End
   Begin VB.Label Label8 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   1800
      TabIndex        =   15
      Top             =   2400
      Width           =   600
   End
   Begin VB.Label Label7 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   7200
      TabIndex        =   14
      Top             =   1800
      Width           =   600
   End
   Begin VB.Label Label6 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   7200
      TabIndex        =   13
      Top             =   2400
      Width           =   600
   End
   Begin VB.Image Frame6 
      Height          =   435
      Left            =   2220
      Picture         =   "frogger.frx":17EB
      Top             =   8400
      Width           =   720
   End
   Begin VB.Image fcar3 
      Height          =   435
      Left            =   5700
      Picture         =   "frogger.frx":19E5
      Top             =   7800
      Width           =   720
   End
   Begin VB.Image framecar2 
      Height          =   435
      Left            =   7380
      Picture         =   "frogger.frx":1BDF
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image Frame1 
      Height          =   435
      Left            =   12660
      Picture         =   "frogger.frx":1DD9
      Top             =   7200
      Width           =   720
   End
   Begin VB.Image car 
      Height          =   435
      Left            =   1800
      Picture         =   "frogger.frx":1FD3
      Top             =   6600
      Width           =   720
   End
   Begin VB.Label Label5 
      BackColor       =   &H00008000&
      Height          =   615
      Left            =   6600
      TabIndex        =   12
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label4 
      BackColor       =   &H00008000&
      Height          =   615
      Left            =   6000
      TabIndex        =   11
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00008000&
      Height          =   615
      Left            =   6000
      TabIndex        =   10
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00008000&
      Height          =   615
      Index           =   0
      Left            =   10200
      TabIndex        =   9
      Top             =   4800
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008000&
      Height          =   615
      Index           =   0
      Left            =   9600
      TabIndex        =   8
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label lbllily5 
      BackColor       =   &H00008000&
      Height          =   615
      Left            =   6000
      TabIndex        =   5
      Top             =   5400
      Width           =   615
   End
   Begin VB.Label lbllily3 
      BackColor       =   &H00008000&
      Height          =   615
      Left            =   10200
      TabIndex        =   4
      Top             =   4200
      Width           =   615
   End
   Begin VB.Label lbllily2 
      BackColor       =   &H00008000&
      Height          =   600
      Left            =   10800
      TabIndex        =   3
      Top             =   5400
      Width           =   600
   End
   Begin VB.Line Line 
      BorderColor     =   &H8000000B&
      Index           =   4
      X1              =   0
      X2              =   13500
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Line Line 
      BorderColor     =   &H8000000B&
      Index           =   3
      X1              =   0
      X2              =   13500
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line 
      BorderColor     =   &H8000000B&
      Index           =   2
      X1              =   0
      X2              =   13500
      Y1              =   7800
      Y2              =   7800
   End
   Begin VB.Line Line 
      BorderColor     =   &H8000000B&
      Index           =   1
      X1              =   0
      X2              =   13500
      Y1              =   8400
      Y2              =   8400
   End
   Begin VB.Line Line 
      BorderColor     =   &H8000000B&
      Index           =   0
      X1              =   0
      X2              =   13500
      Y1              =   9000
      Y2              =   9000
   End
   Begin VB.Label lblroad 
      BackColor       =   &H80000007&
      Height          =   3015
      Left            =   -60
      TabIndex        =   2
      Top             =   6600
      Width           =   13800
   End
   Begin VB.Label lblbw2 
      BackColor       =   &H80000004&
      Height          =   600
      Left            =   12600
      TabIndex        =   1
      Top             =   600
      Width           =   600
   End
   Begin VB.Label lblw2 
      BackColor       =   &H80000004&
      Height          =   600
      Left            =   7800
      TabIndex        =   0
      Top             =   600
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsx As Integer
Dim lsy As Integer
Dim xi As Integer
Dim yi As Integer
Dim arrscore(3000) As Integer
Dim c As Long
Dim i As Long
Dim path As String
Dim fname As String
Dim j As Long
Dim pscore As String
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Sub savescore()
    Framefrog.Left = 6600
    Framefrog.Top = 9600
    Frame2.Visible = True
    cmdleaderboard.Visible = True
    frogtimer.Enabled = False
    cartimer.Enabled = False
Dim ans As String
ans = vbNo
Score = Val(Score)
Do While ans = vbNo
    fname = Trim(UCase(InputBox("Enter a name for your file to save:", "Filename", "Bob", vbOKCancel)))
    path = "c:\" + fname + Score + ".txt"
    If fname = "" Then
        Exit Do
    End If

ans = MsgBox(path, vbYesNo, "Is this the name you want?")
Loop
If ans = vbYes Then
    path = "c:\" + fname + Score + ".txt"
    Open path For Output As #1
        Print #1, fname + ", " + Score
    Close #1
    Frame2.Visible = False
    frogtimer.Enabled = True
ElseIf ans <> vbYes Then
    cmdenter.SetFocus
End If
Frame2.Visible = True
End Sub
Sub startover()
Score = 0
Framefrog.Left = 6600
Framefrog.Top = 9600
frog = LoadPicture("f:\frogger\untitled.gif")
frogtimer.Enabled = True
cartimer.Enabled = True
End Sub

Private Sub cartimer_Timer()

'cars moving at a continuous speed
car.Move car.Left + lsx * 600, car.Top + lsy * 600
Frame1.Move Frame1.Left + lsx * 600, Frame1.Top + lsy * 600
car4.Move car4.Left + lsx * 600, car4.Top + lsy * 600
fcar3.Move fcar3.Left + lsx * 600, fcar3.Top + lsy * 600
Frame6.Move Frame6.Left + lsx * 600, Frame6.Top + lsy * 600
framecar2.Move framecar2.Left + lsx * 600, framecar2.Top + lsy * 600
car2.Move car2.Left + lsx * 600, car2.Top + lsy * 600
car3.Move car3.Left + lsx * 600, car3.Top + lsy * 600
car5.Move car5.Left + lsx * 600, car5.Top + lsy * 600
car6.Move car6.Left + lsx * 600, car6.Top + lsy * 600
car7.Move car7.Left + lsx * 600, car7.Top + lsy * 600
car8.Move car8.Left + lsx * 600, car8.Top + lsy * 600
car9.Move car9.Left + lsx * 600, car9.Top + lsy * 600
car10.Move car10.Left + lsx * 600, car10.Top + lsy * 600
car11.Move car11.Left + lsx * 600, car11.Top + lsy * 600


'car and frog collision
If Framefrog.Top = car8.Top Then
    If Framefrog.Left = car8.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
        frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If

If Framefrog.Top = car11.Top Then
    If Framefrog.Left = car11.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
        frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If

If Framefrog.Top = car9.Top Then
    If Framefrog.Left = car9.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
        frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If

If Framefrog.Top = car10.Top Then
    If Framefrog.Left = car10.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
        frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If


If Framefrog.Top = car.Top Then
    If Framefrog.Left = car.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If
If Framefrog.Top = car2.Top Then
    If Framefrog.Left = car2.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If
If Framefrog.Top = car3.Top Then
    If Framefrog.Left = car3.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If
If Framefrog.Top = car7.Top Then
    If Framefrog.Left = car7.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If

If Framefrog.Top = car6.Top Then
    If Framefrog.Left = car6.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If

If Framefrog.Top = fcar3.Top Then
    If Framefrog.Left = fcar3.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If

If Framefrog.Top = car5.Top Then
    If Framefrog.Left = car5.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If

If Framefrog.Top = Frame6.Top Then
    If Framefrog.Left = Frame6.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
         ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If

If Framefrog.Top = framecar2.Top Then
    If Framefrog.Left = framecar2.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If

If Framefrog.Top = Frame1.Top Then
    If Framefrog.Left = Frame1.Left Then
    frogtimer.Enabled = False
    cartimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                    frogtimer.Enabled = True
                    cartimer.Enabled = True
                Else
                    End
                End If
        End If
    End If
End If



'cars reseting their position
If car.Left + car.Width + 300 > Width + 100 Then
car.Move car.Left = Width
End If

If Frame1.Left + Frame1.Width + 300 > Width + 100 Then
Frame1.Move Frame1.Left = Width
End If

If car4.Left + car4.Width + 300 > Width + 100 Then
car4.Move car4.Left = Width
End If

If fcar3.Left + fcar3.Width + 300 > Width + 100 Then
fcar3.Move fcar3.Left = Width
End If

If Frame6.Left + Frame6.Width + 300 > Width + 100 Then
Frame6.Move Frame6.Left = Width
End If

If car2.Left + car2.Width + 300 > Width + 100 Then
car2.Move car2.Left = Width
End If

If car3.Left + car3.Width + 300 > Width + 100 Then
car3.Move car3.Left = Width
End If

If car5.Left + car5.Width + 300 > Width + 100 Then
car5.Move car5.Left = Width
End If

If car6.Left + car6.Width + 300 > Width + 100 Then
car6.Move car6.Left = Width
End If

If car7.Left + car7.Width + 300 > Width + 100 Then
car7.Move car7.Left = Width
End If

If car8.Left + car8.Width + 300 > Width + 100 Then
car8.Move car8.Left = Width
End If

If car9.Left + car9.Width + 300 > Width + 100 Then
car9.Move car9.Left = Width
End If

If car10.Left + car10.Width + 300 > Width + 100 Then
car10.Move car10.Left = Width
End If

If car11.Left + car11.Width + 300 > Width + 100 Then
car11.Move car11.Left = Width
End If
End Sub

Private Sub cmdenter_Click()
    Score = 0
    Frame2.Visible = False
    frogtimer.Enabled = True
    cmdleaderboard.Enabled = True
    cmdleaderboard.Visible = True
    lstleader.Visible = False
    Framefrog.Left = 6600
    Framefrog.Top = 9600
    cartimer.Enabled = True
End Sub


Private Sub cmdleaderboard_Click()
lstleader.Visible = True
    path = "c:\" + fname + Score + ".txt"
    Open path For Input As #1
        lstleader.AddItem fname + ", " + Score
    Close #1
cmdleaderboard.Enabled = False
End Sub

Private Sub Form_Load()
lsx = 1
lsy = 0
Score = 0

car.Width = 600
Frame1.Width = 600
framecar2.Width = 600
fcar3.Width = 600
car4.Width = 600
Frame6.Width = 600
car2.Width = 600
car3.Width = 600
car5.Width = 600
car6.Width = 600
frog = LoadPicture("f:\frogger\untitled.gif")
frogtimer.Enabled = False

End Sub


Private Sub frogtimer_Timer()

'frog moving
If (GetAsyncKeyState(87)) Then
    Framefrog.Top = Framefrog.Top - 600
    frog = LoadPicture("f:\frogger\untitled.gif")
End If
If (GetAsyncKeyState(65)) Then
    Framefrog.Left = Framefrog.Left - 600
    frog = LoadPicture("f:\frogger\frogleft.gif")
End If
If (GetAsyncKeyState(68)) Then
    Framefrog.Left = Framefrog.Left + 600
    frog = LoadPicture("f:\frogger\frogright.gif")
End If
If (GetAsyncKeyState(83)) Then
    Framefrog.Top = Framefrog.Top + 600
    frog = LoadPicture("f:\frogger\frogdown.gif")
End If

'if framefrog goes off right side of the map
If Framefrog.Left + Framefrog.Width + 300 > Width + 100 Then
    Framefrog.Move Framefrog.Left = Width
End If

'if framefrog goes off left side of the map
If Framefrog.Left < 0 Then
    Framefrog.Left = 12600
End If

'if framefrog goes down
If Framefrog.Top > 9600 Then
    Framefrog.Top = 9600
End If
'first water
If Framefrog.Top = 5400 Then
    If Framefrog.Left = 6000 Or Framefrog.Left = 10800 Then
        Form1.SetFocus
    Else
    frogtimer.Enabled = False
    cartimer.Enabled = False
        frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                Else
                    End
                End If
        End If
    End If
End If

If Framefrog.Top = 4800 Then
    If Framefrog.Left = 6000 Or Framefrog.Left = 10200 Then
        Form1.SetFocus
    Else
    frogtimer.Enabled = False
    cartimer.Enabled = False
        frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying no will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                Else
                    End
                End If
        End If
    End If
End If
If Framefrog.Top = 4200 Then
    If Framefrog.Left = 6000 Or Framefrog.Left = 10200 Then
        Form1.SetFocus
    Else
    frogtimer.Enabled = False
    cartimer.Enabled = False
        frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                Else
                    End
                End If
        End If
    End If
End If
If Framefrog.Top = 3600 Then
    If Framefrog.Left = 6600 Or Framefrog.Left = 9600 Then
                Form1.SetFocus
    Else
    cartimer.Enabled = False
    frogtimer.Enabled = False
        frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                Else
                    End
                End If
        End If
    End If
End If
        
        
'second water
If Framefrog.Top = 2400 Then
    If Framefrog.Left = 7200 Or Framefrog.Left = 3000 Or Framefrog.Left = 12000 Or Framefrog.Left = 1800 Then
            Form1.SetFocus
        Else
        cartimer.Enabled = False
        frogtimer.Enabled = False
        frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                Else
                    End
                End If
        End If
    End If
End If

If Framefrog.Top = 1800 Then
    If Framefrog.Left = 2400 Or Framefrog.Left = 12600 Or Framefrog.Left = 7200 Or Framefrog.Left = 1200 Then
        Form1.SetFocus
    Else
    cartimer.Enabled = False
    frogtimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                Else
                    End
                End If
        End If
    End If
End If
 If Framefrog.Top = 1200 Then
    If Framefrog.Left = 1800 Or Framefrog.Left = 7800 Or Framefrog.Left = 12600 Then
        Form1.SetFocus
    Else
    cartimer.Enabled = False
    frogtimer.Enabled = False
    frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                Else
                    End
                End If
        End If
    End If
End If
If Framefrog.Top = 0 Then
frogtimer.Enabled = False
cartimer.Enabled = False
frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                Else
                    End
                End If
        End If
    End If

'final accomplishment
If Framefrog.Top = 600 Then
    If Framefrog.Left = 7800 Or Framefrog.Left = 12600 Then
        frogtimer.Enabled = False
        cartimer.Enabled = False
        i = MsgBox("Congrats you scored a point, press YES to continue and NO to end game", vbYesNo)
        If i = vbYes Then
            cartimer.Enabled = True
            frogtimer.Enabled = True
            Framefrog.Left = 6600
            Framefrog.Top = 9600
            Score = Score + 1
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                Else
                    End
                End If
        End If
    End If
End If
If Framefrog.Top = 600 Then
    If Framefrog.Left = 1200 Then
        Score = Score + 2
        cartimer.Enabled = False
        frogtimer.Enabled = False
        i = MsgBox("Congrats youscored 2 points, press YES to continue and NO to end game.", vbYesNo)
        If i = vbYes Then
        cartimer.Enabled = True
        frogtimer.Enabled = True
            Framefrog.Left = 6600
            Framefrog.Top = 9600
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                Else
                    End
                End If
        End If
    End If
End If
If Framefrog.Top = 600 Then
    If Framefrog.Left < 1200 Or Framefrog.Left < 7800 Or Framefrog.Left < 12000 Or Framefrog.Left > 12000 Then
        frogtimer.Enabled = False
        cartimer.Enabled = False
        frog = LoadPicture("f:\frogger\dead.gif")
        i = MsgBox("Game Over, would you like to start over, saying NO will end this game?", vbYesNo)
        If i = vbYes Then
            startover
        ElseIf vbNo Then
            j = MsgBox("Do you wish to save your score", vbYesNo)
                If j = vbYes Then
                    savescore
                Else
                    End
                End If
        End If
    End If
End If
End Sub

