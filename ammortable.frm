VERSION 5.00
Begin VB.Form ammortable 
   Caption         =   "Ammort Table"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   15165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame 
      Caption         =   "Saved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   2
      Left            =   6240
      TabIndex        =   28
      Top             =   2940
      Width           =   5895
      Begin VB.Label lblsaveprice 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3180
         TabIndex        =   31
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label 
         Caption         =   "Months and"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   1260
         TabIndex        =   30
         Top             =   600
         Width           =   1875
      End
      Begin VB.Label lblmonths 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   120
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Amortization Table"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Index           =   3
      Left            =   180
      TabIndex        =   14
      Top             =   5100
      Width           =   13095
      Begin VB.HScrollBar hsbPayment 
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   1380
         Width           =   12975
      End
      Begin VB.TextBox txtamorttable 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   660
         Width           =   12915
      End
      Begin VB.Label Label 
         Caption         =   "Monthly Principal:"
         Height          =   255
         Index           =   10
         Left            =   11340
         TabIndex        =   21
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label 
         Caption         =   "Monthly Interest:"
         Height          =   255
         Index           =   9
         Left            =   8880
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label 
         Caption         =   "Total Interest:"
         Height          =   195
         Index           =   8
         Left            =   6540
         TabIndex        =   19
         Top             =   420
         Width           =   1095
      End
      Begin VB.Label Label 
         Caption         =   "Current Amount:"
         Height          =   195
         Index           =   7
         Left            =   3120
         TabIndex        =   18
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Year Number:"
         Height          =   195
         Index           =   6
         Left            =   1680
         TabIndex        =   17
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label 
         Caption         =   "Payment Number:"
         Height          =   255
         Index           =   5
         Left            =   180
         TabIndex        =   16
         Top             =   420
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdquit 
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7380
      TabIndex        =   13
      Top             =   4260
      Width           =   1935
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4860
      TabIndex        =   12
      Top             =   4260
      Width           =   2235
   End
   Begin VB.CommandButton cmdcalc 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1800
      TabIndex        =   11
      Top             =   4260
      Width           =   2955
   End
   Begin VB.Frame Frame 
      Caption         =   "Monthly Payment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Index           =   1
      Left            =   1800
      TabIndex        =   9
      Top             =   2880
      Width           =   4035
      Begin VB.Label lblmonthlypayment 
         BorderStyle     =   1  'Fixed Single
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
         Left            =   420
         TabIndex        =   10
         Top             =   540
         Width           =   3135
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Enter Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   14895
      Begin VB.TextBox txtonetime 
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
         Left            =   12780
         TabIndex        =   27
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox Txtyrlypay 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9960
         TabIndex        =   25
         Top             =   1260
         Width           =   2655
      End
      Begin VB.TextBox txtxtra 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6780
         TabIndex        =   8
         Text            =   "100"
         Top             =   1260
         Width           =   2775
      End
      Begin VB.TextBox txtyears 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4740
         TabIndex        =   6
         Text            =   "30"
         Top             =   1200
         Width           =   1875
      End
      Begin VB.TextBox Txtyrlyrate 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2580
         TabIndex        =   4
         Text            =   "12"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtloanamount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   180
         TabIndex        =   2
         Text            =   "100000"
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label 
         Caption         =   "Extra One Time:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   12
         Left            =   12720
         TabIndex        =   26
         Top             =   900
         Width           =   2115
      End
      Begin VB.Label Label 
         Caption         =   "Extra Yearly Payment:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   11
         Left            =   9840
         TabIndex        =   24
         Top             =   900
         Width           =   3135
      End
      Begin VB.Label Label 
         Caption         =   "Extra Monthly Payment:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   6780
         TabIndex        =   7
         Top             =   900
         Width           =   3135
      End
      Begin VB.Label Label 
         Caption         =   "Years:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   4740
         TabIndex        =   5
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label 
         Caption         =   "Yearly Interest:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   2520
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label 
         Caption         =   "Loan Amount:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   2115
      End
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Loan Amortization"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   4
      Left            =   2040
      TabIndex        =   23
      Top             =   0
      Width           =   11475
   End
End
Attribute VB_Name = "ammortable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim amorttable(360) As String

Private Sub cmdcalc_Click()

Dim loanamount As Currency, monthlypayment, saveprice, onetime, xtrayear, totalp, nrmltotalint As Currency
Dim yrlyrate As Single, monthlyrate As Single
Dim years As Integer, payments As Integer, xtra As Integer, months As Integer
Dim paymentnumber As Integer
Dim monthlyint, monthlyint2 As Currency
Dim totalint As Currency, currentamt As Currency, monthlyprin As Currency
Dim yearnumber As Integer, year As Integer, displine As String

    xtra = Val(txtxtra)
    onetime = Val(txtonetime)
    xtrayear = Val(Txtyrlypay)
    loanamount = Val(txtloanamount)
    yrlyrate = Val(Txtyrlyrate)
    years = Val(txtyears)
    monthlyrate = yrlyrate / 1200
    payments = years * 12
    

If txtloanamount = "" Or Txtyrlyrate = "" Or txtyears = "" Then


    Error.Visible = True
    cmdclear.SetFocus
Else
    
    monthlypayment = (loanamount * monthlyrate / (1 - (1 + monthlyrate) ^ (-payments))) + xtra
    totalp = ((loanamount * monthlyrate / (1 - (1 + monthlyrate) ^ -payments)) * payments)
    nrmltotalint = totalp - loanamount


    lblmonthlypayment = Format$(monthlypayment, "Currency")
    
    totalint = 0
    currentamt = loanamount


    For paymentnumber = 1 To payments
        If paymentnumber Mod 12 = 0 Then
            currentamt = currentamt - xtrayear
        Else
            monthlypayment = monthlypayment
        End If
        If paymentnumber = 1 Then
            currentamt = currentamt - onetime
        End If
        
        saveprice = (payments * (monthlypayment - xtra)) - (paymentnumber * (monthlypayment))
        months = payments - paymentnumber
        monthlyint = currentamt * monthlyrate
        totalint = totalint + monthlyint
        
        
            If paymentnumber Mod 12 = 0 Then
                year = paymentnumber / 12
            Else
                year = Int(paymentnumber / 12) + 1
            End If
        
        If currentamt < monthlypayment Then
            
            monthlyprin = currentamt
            currentamt = 0
            displine = vbTab + Format$(paymentnumber, "####")
            displine = displine + vbTab + Format$(year, "#0")
            displine = displine + vbTab + Format$(currentamt, "Currency")
            displine = displine + vbTab + vbTab + Format$(totalint, "Currency")
            displine = displine + vbTab + vbTab + Format$(monthlyint, "Currency")
            displine = displine + vbTab + vbTab + Format$(monthlyprin, "Currency")


            amorttable(paymentnumber) = displine
            lblmonths = months
            saveprice = nrmltotalint - totalint
            lblsaveprice = Format$(saveprice, "currency")
            
            
            
            
            Exit For
        Else
            currentamt = currentamt + monthlyint - monthlypayment
            monthlyprin = monthlypayment - monthlyint
            displine = vbTab + Format$(paymentnumber, "####")
            displine = displine + vbTab + Format$(year, "#0")
            displine = displine + vbTab + Format$(currentamt, "Currency")
            displine = displine + vbTab + vbTab + Format$(totalint, "Currency")
            displine = displine + vbTab + vbTab + Format$(monthlyint, "Currency")
            displine = displine + vbTab + vbTab + Format$(monthlyprin, "Currency")


            amorttable(paymentnumber) = displine
            lblmonths = months
            saveprice = nrmltotalint - totalint
            lblsaveprice = Format$(saveprice, "currency")
        End If

    Next paymentnumber


            
    hsbPayment.Min = 1
    hsbPayment.Max = paymentnumber
    hsbPayment.LargeChange = 12
    hsbPayment.Value = 1
    txtamorttable = amorttable(1)

End If

End Sub
Private Sub cmdclear_Click()

lblmonthlypayment = ""
Txtyrlyrate = ""
txtamorttable = ""
txtyears = ""
txtxtra = ""
txtloanamount = ""
lblmonths = ""
lblsaveprice = ""
Error.Visible = False
txtloanamount.SetFocus

End Sub

Private Sub cmdquit_Click()

End

End Sub

Private Sub Form_Load()

End Sub

Private Sub hsbPayment_Change()

txtamorttable = amorttable(hsbPayment.Value)

End Sub



