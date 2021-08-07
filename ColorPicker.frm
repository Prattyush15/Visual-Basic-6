VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form1"
   ClientHeight    =   5340
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   ScaleHeight     =   5340
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdback 
      Caption         =   "Go back to drawing"
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   60
      Width           =   2595
   End
   Begin VB.HScrollBar hsb 
      Height          =   315
      Left            =   180
      TabIndex        =   5
      Top             =   4620
      Width           =   5655
   End
   Begin VB.HScrollBar hsg 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   2700
      Width           =   5655
   End
   Begin VB.HScrollBar hsr 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1140
      Width           =   5655
   End
   Begin VB.TextBox txtg 
      Height          =   855
      Left            =   180
      TabIndex        =   2
      Text            =   "txtg"
      Top             =   1620
      Width           =   1515
   End
   Begin VB.TextBox txtb 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Text            =   "txtb"
      Top             =   3540
      Width           =   1515
   End
   Begin VB.TextBox txtr 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Text            =   "txtr"
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim r(255) As String, g(255) As String, b(255) As String
Dim rr As Integer, gg As Integer, bb As Integer


Private Sub cmdblack_Click()
Paint.Show
colorpicker.Hide
'Paint.Label.BackColor = RGB(r(hsr.Value), g(hsg.value), b (hsb.Value))
End Sub

Private Sub cmdback_Click()
Paint.Show
colorpicker.Hide
End Sub

Private Sub hsr_Change()
txtr = r(hsr.Value)
rr = txtr.Text
lblcolor.BackColor = RGB(r(hsr.Value), g(hsg.Value), b(hsb.Value))
End Sub
Private Sub hsg_Change()
txtg = g(hsg.Value)
gg = txtg.Text
lblcolor.BackColor = RGB(r(hsr.Value), g(hsg.Value), b(hsb.Value))
End Sub
Private Sub hsb_Change()
txtb = b(hsb.Value)
bb = txtb.Text
lblcolor.BackColor = RGB(r(hsr.Value), g(hsg.Value), b(hsb.Value))
End Sub
Private Sub Form_Load()
Dim i As Integer
Dim displine As Integer
For i = 0 To 255
    r(i) = i
    g(i) = 1
    b(i) = i
Next i

hsr.Min = 0
hsr.Max = 255
hsr.LargeChange = 12

hsg.Min = 0
hsg.Max = 255
hsg.LargeChange = 12

hs.Min = 0
hsb.Max = 255
hsb.LargeChange = 12

txtr = r(0)
txtg = g(0)
txtb = b(0)
End Sub

