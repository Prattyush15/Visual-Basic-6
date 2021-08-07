VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   ScaleHeight     =   8040
   ScaleWidth      =   14220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdgl 
      Caption         =   "Find GCF and LCM"
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
      Left            =   11880
      TabIndex        =   31
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox txtnumberoffactors2 
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
      Left            =   11700
      ScrollBars      =   1  'Horizontal
      TabIndex        =   22
      Top             =   4020
      Width           =   2355
   End
   Begin VB.TextBox txtsum2 
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
      Left            =   9600
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   21
      Top             =   4080
      Width           =   1875
   End
   Begin VB.TextBox txtNumberofFactors1 
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
      Left            =   11340
      ScrollBars      =   1  'Horizontal
      TabIndex        =   20
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox txtsum1 
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
      Left            =   8640
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   19
      Top             =   1740
      Width           =   2235
   End
   Begin VB.ListBox lstfactors2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      ItemData        =   "Factor.frx":0000
      Left            =   6960
      List            =   "Factor.frx":0002
      TabIndex        =   18
      Top             =   3660
      Width           =   2475
   End
   Begin VB.ListBox lstfactors1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      ItemData        =   "Factor.frx":0004
      Left            =   5820
      List            =   "Factor.frx":0006
      TabIndex        =   17
      Top             =   1320
      Width           =   2595
   End
   Begin VB.TextBox txtlcm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   7020
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   6180
      Width           =   2955
   End
   Begin VB.TextBox txtgcf 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   6180
      Width           =   3255
   End
   Begin VB.TextBox txtfactors2 
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
      Left            =   3300
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   11
      Top             =   3960
      Width           =   3375
   End
   Begin VB.TextBox txtnumber2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   10
      Top             =   3960
      Width           =   3075
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
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
      Left            =   11940
      TabIndex        =   6
      Top             =   7500
      Width           =   2235
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
      Height          =   975
      Left            =   12000
      TabIndex        =   5
      Top             =   6540
      Width           =   2175
   End
   Begin VB.TextBox txtfactors1 
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
      Left            =   2220
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   4
      Top             =   1680
      Width           =   3435
   End
   Begin VB.TextBox txtnumber1 
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
      Left            =   60
      TabIndex        =   2
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Divisor List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   16
      Left            =   6900
      TabIndex        =   30
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Divisors List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   15
      Left            =   5820
      TabIndex        =   29
      Top             =   900
      Width           =   2535
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "LCM"
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
      Index           =   14
      Left            =   6960
      TabIndex        =   28
      Top             =   5700
      Width           =   3135
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "GCF"
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
      Index           =   13
      Left            =   3180
      TabIndex        =   27
      Top             =   5700
      Width           =   3255
   End
   Begin VB.Label Label 
      Caption         =   "Number of Factors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   12
      Left            =   11580
      TabIndex        =   26
      Top             =   3600
      Width           =   3075
   End
   Begin VB.Label Label 
      Caption         =   "Sum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   11
      Left            =   9840
      TabIndex        =   25
      Top             =   3600
      Width           =   1035
   End
   Begin VB.Label Label 
      Caption         =   "Number of Factors"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   10
      Left            =   11100
      TabIndex        =   24
      Top             =   1380
      Width           =   2835
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Sum"
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
      Index           =   9
      Left            =   8520
      TabIndex        =   23
      Top             =   1320
      Width           =   2355
   End
   Begin VB.Label lblnotprime2 
      BackColor       =   &H000000FF&
      Caption         =   "It is not Prime!"
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
      Left            =   120
      TabIndex        =   14
      Top             =   5040
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Label lblprime2 
      BackColor       =   &H0000FF00&
      Caption         =   "It is Prime!"
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
      Left            =   180
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Divisors"
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
      Index           =   5
      Left            =   3660
      TabIndex        =   12
      Top             =   3420
      Width           =   2775
   End
   Begin VB.Label Label 
      Caption         =   "Enter 2nd Number"
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
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Width           =   3315
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblnotprime1 
      BackColor       =   &H000000FF&
      Caption         =   "It is not Prime!"
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
      Left            =   60
      TabIndex        =   8
      Top             =   2460
      Visible         =   0   'False
      Width           =   2595
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPrime1 
      BackColor       =   &H0000C000&
      Caption         =   "It is Prime!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   2460
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Divisors Multi-Line"
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
      Index           =   2
      Left            =   1440
      TabIndex        =   3
      Top             =   1140
      Width           =   5055
   End
   Begin VB.Label Label 
      Caption         =   "Enter a Number"
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
      Left            =   60
      TabIndex        =   1
      Top             =   840
      Width           =   1635
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Prime numbers and Factors"
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
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   60
      Width           =   13215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Number1 As Long
Dim Number2 As Long
Dim Number3 As Long
Dim Number1Check As Boolean
Dim Number2Check As Boolean
Dim Number3Check As Boolean

Sub CheckBoth()
Dim i As Integer
Dim GCF As Long
Dim LCM As Long

If Number1Check = True And Number2Check = True And Number3Check = False Then
    For i = 1 To Number1
        If Number1 Mod i = 0 And Number2 Mod i = 0 And Number3 Mod i = 0 Then
            GCF = i
        End If
    Next i
    txtgcf.Text = GCF
    
    LCM = (Number1 / GCF) * Number2
    txtlcm.Text = LCM
End If
End Sub


Private Sub cmdclear_Click()
txtnumber1.Text = ""
txtfactors1.Text = ""
lblPrime1.Visible = False
lblnotprime1.Visible = False
txtnumber2.Text = ""
txtfactors2.Text = ""
lblprime2.Visible = False
lblnotprime2.Visible = False



lstfactors1.Clear
lstfactors2.Clear

txtNumberofFactors1 = ""
txtnumberoffactors2 = ""

txtsum1 = ""
txtsum2 = ""

txtgcf = ""
txtlcm = ""
txtnumber1.Enabled = True
txtnumber2.Enabled = True

txtnumber1.SetFocus
Number1Check = False
Number2Check = False

End Sub

Private Sub cmdexit_Click()
End
End Sub


Private Sub cmdgl_Click()
CheckBoth
End Sub

Private Sub txtNumber1_KeyPress(KeyAscii As Integer)
Dim i As Long
Dim Sum1 As Long
Dim Factors1 As String
Dim NumberOfFactors1 As Long
If KeyAscii = 13 Then
    If Not IsNumeric(txtnumber1) Then
        MsgBox ("Please enter a valid number.")
    Else
        Number1 = Val(txtnumber1)
        
        For i = 1 To Number1
            If Number1 Mod i = 0 Then
                Factors1 = Factors1 + Str(i) + ""
                lstfactors1.AddItem (Str(i))
                Sum1 = Sum1 + i
                NumberOfFactors1 = NumberOfFactors1 + 1
            End If
        Next i
        txtfactors1.Text = Factors1
        txtsum1.Text = Sum1
        txtNumberofFactors1.Text = NumberOfFactors1
        If Sum1 = Number1 + 1 Then
            lblPrime1.Visible = True
        Else
            lblnotprime1.Visible = True
        End If
        Number1Check = True
        txtnumber2.SetFocus
        
    End If
End If
End Sub

Private Sub txtNumber2_KeyPress(KeyAscii As Integer)
Dim i As Long
Dim Sum2 As Long
Dim Factors2 As String
Dim NumberOfFactors2 As Long

If KeyAscii = 13 Then
    If Not IsNumeric(txtnumber2) Then
        MsgBox ("Please enter a valid number.")
    Else
        Number2 = Val(txtnumber2)
        
        For i = 1 To Number2
            If Number2 Mod i = 0 Then
                Factors2 = Factors2 + Str(i) + ""
                lstfactors2.AddItem (Str(i))
                Sum2 = Sum2 + i
                NumberOfFactors2 = NumberOfFactors2 + 1
            End If
        Next i
        txtfactors2.Text = Factors2
        txtsum2.Text = Sum2
        txtnumberoffactors2.Text = NumberOfFactors2
        If Sum2 = Number2 + 1 Then
            lblprime2.Visible = True
        Else
            lblnotprime2.Visible = True
        End If
        Number2Check = True
        txtnumber2.SetFocus
        
        
    End If
End If
End Sub


