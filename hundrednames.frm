VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14580
   LinkTopic       =   "Form1"
   ScaleHeight     =   9450
   ScaleWidth      =   14580
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdend 
      Caption         =   "End"
      Height          =   495
      Left            =   12480
      TabIndex        =   15
      Top             =   3240
      Width           =   1515
   End
   Begin VB.CommandButton cmdclearfull 
      Caption         =   "Clear the hundred random names"
      Height          =   495
      Left            =   12360
      TabIndex        =   14
      Top             =   2220
      Width           =   1755
   End
   Begin VB.CommandButton cmdfullclear 
      Caption         =   "Clear All Listboxes"
      Height          =   555
      Left            =   12180
      TabIndex        =   13
      Top             =   1260
      Width           =   2055
   End
   Begin VB.ListBox lstfull 
      Height          =   6495
      ItemData        =   "hundrednames.frx":0000
      Left            =   9300
      List            =   "hundrednames.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   12
      Top             =   1080
      Width           =   2595
   End
   Begin VB.CommandButton cmdcreate 
      Caption         =   "Create A Hundred Full Names"
      Height          =   615
      Left            =   9600
      TabIndex        =   11
      Top             =   8340
      Width           =   2175
   End
   Begin VB.CommandButton cmdmn 
      Caption         =   "Fill Male First Names"
      Height          =   615
      Left            =   6600
      TabIndex        =   10
      Top             =   8340
      Width           =   2175
   End
   Begin VB.CommandButton cmdfn 
      Caption         =   "Fill Female First Names"
      Height          =   615
      Left            =   3540
      TabIndex        =   9
      Top             =   8340
      Width           =   2235
   End
   Begin VB.CommandButton cmdfill 
      Caption         =   "Fill Last Names"
      Height          =   615
      Left            =   480
      TabIndex        =   8
      Top             =   8340
      Width           =   2355
   End
   Begin VB.ListBox lstmalefirst 
      Height          =   6495
      ItemData        =   "hundrednames.frx":0004
      Left            =   6480
      List            =   "hundrednames.frx":0006
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ListBox lstfemalefirst 
      Height          =   6495
      ItemData        =   "hundrednames.frx":0008
      Left            =   3480
      List            =   "hundrednames.frx":000A
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ListBox lstlastname 
      Height          =   6495
      ItemData        =   "hundrednames.frx":000C
      Left            =   480
      List            =   "hundrednames.frx":000E
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label 
      Caption         =   "Hundred Full Names"
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
      Index           =   4
      Left            =   9300
      TabIndex        =   19
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label Label 
      Caption         =   "Male First"
      Height          =   435
      Index           =   3
      Left            =   6480
      TabIndex        =   18
      Top             =   600
      Width           =   2355
   End
   Begin VB.Label Label 
      Caption         =   "Female First"
      Height          =   435
      Index           =   2
      Left            =   3540
      TabIndex        =   17
      Top             =   660
      Width           =   2355
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Last Names"
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
      Index           =   1
      Left            =   480
      TabIndex        =   16
      Top             =   600
      Width           =   2415
   End
   Begin VB.Label lblcountfull 
      BorderStyle     =   1  'Fixed Single
      Height          =   435
      Left            =   9480
      TabIndex        =   7
      Top             =   7620
      Width           =   2415
   End
   Begin VB.Label lblcountmn 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   6540
      TabIndex        =   6
      Top             =   7620
      Width           =   2355
   End
   Begin VB.Label lblcountfn 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   3480
      TabIndex        =   5
      Top             =   7620
      Width           =   2415
   End
   Begin VB.Label lblcountl 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   480
      TabIndex        =   4
      Top             =   7620
      Width           =   2355
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "100 Names"
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
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14835
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim path As String
Dim fname As String
Dim filepath As String
Dim strline As String
Dim rnum As Long

Private Sub cmdclearfull_Click()
lstfull.Clear
lblcountfull = ""
End Sub

Private Sub cmdcreate_Click()
Dim count As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim c As Long
Dim a As Long

count = 0
a = Val(lstfull.ListCount)
For a = 1 To 100
Randomize
i = Int(Rnd * lstlastname.ListCount)
j = Int(Rnd * lstfemalefirst.ListCount)
k = Int(Rnd * lstmalefirst.ListCount)
c = Int(Rnd * 2)
If c = 0 Then
lstfull.AddItem (lstlastname.List(i) + ", " + lstmalefirst.List(k))
ElseIf c = 1 Then
lstfull.AddItem (lstlastname.List(i) + ", " + lstfemalefirst.List(j))
End If
count = count + 1
lblcountfull = count
Next
End Sub

Private Sub cmdend_Click()
End
End Sub

Private Sub cmdfill_Click()
Dim count As Long
path = "C:\" + "LastNameCut" + ".txt"
count = 0
Open path For Input As #1
Do While Not EOF(1)
Line Input #1, strline
lstlastname.AddItem strline
count = count + 1
Loop
Close #1
lblcountl = count
If lstlastname.ListCount And lstfemalefirst.ListCount And lstmalefirst.ListCount Then
    cmdcreate.Enabled = True
Else
    cmdcreate.Enabled = False
End If
End Sub

Private Sub cmdfullclear_Click()
lstlastname.Clear
lstfemalefirst.Clear
lstmalefirst.Clear
lstfull.Clear
lblcountl = ""
lblcountfn = ""
lblcountmn = ""
lblcountfull = ""
cmdcreate.Enabled = False
End Sub

Private Sub cmdmn_Click()
Dim count As Long
path = "C:\" + "MaleNamesDictionary" + ".txt"
count = 0
Open path For Input As #1
Do While Not EOF(1)
Line Input #1, strline
lstmalefirst.AddItem strline
count = count + 1
lblcountmn = count
Loop
Close #1
If lstlastname.ListCount And lstfemalefirst.ListCount And lstmalefirst.ListCount Then
    cmdcreate.Enabled = True
Else
    cmdcreate.Enabled = False
End If
End Sub

Private Sub cmdfn_Click()
Dim count As Long
path = "C:\" + "FemaleNamesDictionary" + ".txt"
count = 0
Open path For Input As #1
Do While Not EOF(1)
Line Input #1, strline
lstfemalefirst.AddItem strline
count = count + 1
lblcountfn = count
Loop
Close #1
If lstlastname.ListCount And lstfemalefirst.ListCount And lstmalefirst.ListCount Then
    cmdcreate.Enabled = True
Else
    cmdcreate.Enabled = False
End If
End Sub

Private Sub Form_Load()
cmdcreate.Enabled = False
End Sub
