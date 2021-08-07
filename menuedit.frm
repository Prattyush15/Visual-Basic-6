VERSION 5.00
Begin VB.Form mnuedit 
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14640
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   14640
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtedit 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   14535
   End
   Begin VB.Label lblrnum 
      Height          =   1155
      Left            =   10020
      TabIndex        =   1
      Top             =   360
      Width           =   3975
   End
   Begin VB.Menu menufile 
      Caption         =   "File"
      Begin VB.Menu mmnunew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuopen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnusaveas 
         Caption         =   "Save As"
      End
      Begin VB.Menu menuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Menuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnucopy 
         Caption         =   "Copy"
      End
      Begin VB.Menu mnupaste 
         Caption         =   "Paste"
      End
      Begin VB.Menu mnucut 
         Caption         =   "Cut"
      End
      Begin VB.Menu mnuall 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnudeleteall 
         Caption         =   "Delete All"
      End
   End
   Begin VB.Menu menuoption 
      Caption         =   "Option"
      Begin VB.Menu mnufont 
         Caption         =   "Font"
         Begin VB.Menu mnuarial 
            Caption         =   "Arial"
         End
         Begin VB.Menu mnucomic 
            Caption         =   "Comic Sans"
         End
         Begin VB.Menu mnusans 
            Caption         =   "Sans Serif"
         End
         Begin VB.Menu mnutimes 
            Caption         =   "Times New Roman"
         End
      End
      Begin VB.Menu mnucolor 
         Caption         =   "Background Color"
         Begin VB.Menu mnured 
            Caption         =   "Red"
         End
         Begin VB.Menu mnuyellow 
            Caption         =   "Yellow"
         End
         Begin VB.Menu mnugreen 
            Caption         =   "Green"
         End
         Begin VB.Menu mnuwhite 
            Caption         =   "White"
         End
      End
      Begin VB.Menu menufontsize 
         Caption         =   "Font Size"
         Begin VB.Menu mnu8 
            Caption         =   "8"
         End
         Begin VB.Menu mnu12 
            Caption         =   "12"
         End
         Begin VB.Menu mnu18 
            Caption         =   "18"
         End
         Begin VB.Menu mnu48 
            Caption         =   "48"
         End
      End
   End
End
Attribute VB_Name = "mnuedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BodyText As String
Dim OldBodyText As String
Dim UserFont As String
Dim FontColor As String
Dim BackgroundColor As String
Dim UserFontSize As String
Dim FilePath As String
Dim FormatPath As String
Dim fname As String
Dim path As String
Sub newfile()
Dim Answer As String
Answer = MsgBox("Are you sure you want to create a new document? All unsaved changes will be lost.", vbOKCancel)
    If Answer = vbOK Then
        txtedit.Text = ""
        BodyText = txtedit.Text
        OldBodyText = ""
        FilePath = ""
        UserFont = "Arial Narrow"
        FontColor = vbBlack
        UserFontSize = "12"
        BackgroundColor = vbWhite
        txtedit.Font = "Arial Narrow"
        txtedit.FontSize = "12"
        txtedit.ForeColor = vbBlack
        txtedit.BackColor = vbWhite
        FilePath = ""
        path = ""
    ElseIf Answer = vbCancel Then
        txtedit.SetFocus
    End If
End Sub

Sub s()
FilePath = path
End Sub
Sub SaveFile()
s
If FilePath = "" Then
    SaveAsFile
Else
    Open FilePath For Output As #1
        Print #1, txtedit
    Close #1
End If
End Sub
Sub SaveAsFile()
Dim ans As String
ans = vbNo
s
Do While ans = vbNo
    fname = Trim(UCase(InputBox("Enter a name for your file to save:", "Filename", "Bob", vbOKCancel)))
    path = "c:\" + fname + ".txt"
    If fname = "" Then
        Exit Do
    End If
    
ans = MsgBox(path, vbYesNo, "Is this the name you want?")
Loop
If ans = vbYes Then
    path = "c:\" + fname + ".txt"
    Open path For Output As #1
        Print #1, txtedit
    Close #1
ElseIf ans <> vbYes Then
    txtedit.SetFocus
End If
End Sub

Private Sub Form_Load()

End Sub

Private Sub menuexit_Click()
Dim i As String
i = vbNo
Do While i = vbNo
i = MsgBox("You are about to exit. Do you wish to save?", vbYesNoCancel)
If i = vbYes Then
    SaveFile
    End
ElseIf i = vbNo Then
    End
ElseIf i = vbCancel Then
    txtedit.SetFocus
End If
Loop
End Sub

Private Sub mnu10_Click()
txtedit.FontSize = 10
End Sub
Private Sub mmnunew_Click()

Dim i As String

i = MsgBox("Do you want to save before starting a new document?", vbYesNoCancel)
If i = vbYes Then
    SaveFile
ElseIf i = vbNo Then
    newfile
ElseIf i = vbCancel Then
    txtedit.SetFocus
End If
End Sub

Private Sub mnu12_Click()
txtedit.FontSize = 12
End Sub

Private Sub mnu18_Click()
txtedit.FontSize = 18
End Sub

Private Sub mnu48_Click()
txtedit.FontSize = 48
End Sub

Private Sub mnu8_Click()
txtedit.FontSize = 8
End Sub

Private Sub mnuall_Click()
txtedit.SelStart = 0
txtedit.SelLength = Len(txtedit.Text)
End Sub

Private Sub mnuarial_Click()
UserFont = "Arial"
End Sub

Private Sub mnucomic_Click()
txtedit.Font = "MS Comic Sans"
End Sub

Private Sub mnucopy_Click()
Clipboard.SetText txtedit.SelText
mnupaste.Enabled = True
End Sub

Private Sub mnucut_Click()
Clipboard.SetText txtedit.SelText
txtedit.SelText = ""
mnupaste.Enabled = True
End Sub


Private Sub mnudeleteall_Click()
Dim i As String
i = MsgBox("This process will delete everything you have written and will not save. Do you wish to continue?", vbYesNo)
If i = vbYes Then
    txtedit = ""
ElseIf i = vbNo Then
    txtedit.SetFocus
End If
End Sub

Private Sub mnugreen_Click()
txtedit.BackColor = vbGreen
End Sub
Private Sub mnuopen_Click()
Dim ans As String
Dim i As String
ans = vbNo
i = MsgBox("Do you want to save before opening a new file?", vbYesNo)
If i = vbYes Then
    SaveFile
End If
Do While ans = vbNo
    fname = Trim(UCase(InputBox("Enter the name of the file you wish to open:", "Filename", "Bob")))
    path = "c:\" + fname + ".txt"
    If fname = "" Then
        Exit Do
    End If
    ans = MsgBox(path, vbYesNo, "Is this the name of your file?")
Loop
If ans = vbYes Then
On Error GoTo myerrhandler
    Open path For Input As #1
    Dim filesize As Integer
    filesize = LOF(1)
    txtedit = Input(filesize, #1)
    Close #1
    Exit Sub
If ans = vbNo Then
    txtedit.SetFocus
End If
myerrhandler:
            MsgBox Err.Description
            Exit Sub
End If
End Sub

Private Sub mnupaste_Click()
txtedit.SelText = Clipboard.GetText()
End Sub
Private Sub mnured_Click()
txtedit.BackColor = vbRed
End Sub
Private Sub mnusans_Click()
txtedit.Font = "MS Sans Serif"
End Sub
Private Sub mnusave_Click()
Dim Answer As String
Answer = MsgBox("Are you sure you want to save this file?", vbYesNo)
If Answer = vbYes Then
    SaveFile
ElseIf Answer = vbNo Then
    txtedit.SetFocus
End If
End Sub
Private Sub mnusaveas_Click()
SaveAsFile
End Sub
Private Sub mnutimes_Click()
txtedit.Font = "Times New Roman"
UserFont = "Times New Roman"
End Sub
Private Sub mnuwhite_Click()
txtedit.BackColor = vbWhite
End Sub
Private Sub mnuyellow_Click()
txtedit.BackColor = vbYellow
End Sub

