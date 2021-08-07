VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13395
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   13395
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   7
      Top             =   4320
      Width           =   3615
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9540
      TabIndex        =   6
      Top             =   3000
      Width           =   3735
   End
   Begin VB.CommandButton cmdFindPay 
      Caption         =   "Find Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9480
      TabIndex        =   5
      Top             =   1800
      Width           =   3795
   End
   Begin VB.Frame Frame1 
      Caption         =   "Calculated Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   5580
      TabIndex        =   11
      Top             =   1320
      Width           =   3435
      Begin VB.Label lblTotalPayback 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   17
         Top             =   4020
         Width           =   3375
      End
      Begin VB.Label lblTotalInt 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   16
         Top             =   2580
         Width           =   3375
      End
      Begin VB.Label lblMonthlyPayment 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   60
         TabIndex        =   15
         Top             =   960
         Width           =   3375
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Monthly Payment"
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
         Left            =   60
         TabIndex        =   14
         Top             =   540
         Width           =   3315
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Total Intrest Paid"
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
         Left            =   60
         TabIndex        =   13
         Top             =   2160
         Width           =   3315
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Total Payback"
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
         Left            =   60
         TabIndex        =   12
         Top             =   3480
         Width           =   3315
      End
   End
   Begin VB.Frame Frame 
      Caption         =   "Enter Values"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   420
      TabIndex        =   0
      Top             =   1380
      Width           =   3435
      Begin VB.TextBox txtYrlyRate 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   60
         TabIndex        =   3
         Top             =   2220
         Width           =   3315
      End
      Begin VB.TextBox txtyears 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   3840
         Width           =   3195
      End
      Begin VB.TextBox txtLoanAmount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   60
         TabIndex        =   1
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Years"
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
         Left            =   60
         TabIndex        =   9
         Top             =   3420
         Width           =   3315
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Yearly Intrest Rate"
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
         Left            =   60
         TabIndex        =   8
         Top             =   1800
         Width           =   3315
      End
      Begin VB.Label label1 
         Alignment       =   2  'Center
         Caption         =   "Loan Amount"
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
         Left            =   60
         TabIndex        =   2
         Top             =   540
         Width           =   3315
      End
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Car Loans"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   60
      TabIndex        =   10
      Top             =   0
      Width           =   13695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClear_Click()
'-Clearing the Display Labels
lblMonthlyPayment = ""
lblTotalInt = ""
lblTotalPayback = ""
txtLoanAmount = ""
txtYrlyRate = ""
txtyears = ""
'-Set focus back to the Loan Amount
txtLoanAmount.SetFocus
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdFindPay_Click()
Dim LoanAmount As Currency, MonthlyPayment As Currency
Dim TotalInt As Currency, TotalPayback As Currency
Dim YrlyRate As Single, MonthlyRate As Single
Dim Years As Integer, Payments As Integer
'-Readomg values from the form
LoanAmount = Val(txtLoanAmount)
YrlyRate = Val(txtYrlyRate)
Years = Val(txtyears)
'-Intermediate calculations
MonthlyRate = YrlyRate / 1200
Payments = Years * 12
'-Monthly payment
MonthlyPayment = LoanAmount * MonthlyRate / (1 - (1 + MonthlyRate) ^ (-Payments))
'-Total Payback
TotalPayback = MonthlyPayment * Payments
'-Total Intrest Paid
TotalInt = TotalPayback - LoanAmount
'-Display results
lblMonthlyPayment = Format$(MonthlyPayment, "Currency")
lblTotalInt = Format$(TotalInt, "Currency")
lblTotalPayback = Format$(TotalPayback, "Currency")
cmdClear.SetFocus
End Sub

