VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "Modified Needleman-Wunsch Pattern Matching Technique - Shape #1"
   ClientHeight    =   6960
   ClientLeft      =   285
   ClientTop       =   570
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   7995
   Begin VB.CommandButton Command6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save Image"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3000
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Find Angles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Draw"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      MaskColor       =   &H00C00000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      Enabled         =   0   'False
      Height          =   6135
      Left            =   120
      ScaleHeight     =   405
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   405
      TabIndex        =   0
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   6135
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Enabled = False
Form4.Show
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
Command5.Enabled = True
End Sub

Private Sub Command2_Click()
Dim k As Integer
For k = 0 To N - 1
    Picture1.Line (Arr(k, 0), Arr(k, 1))-(Arr(k + 1, 0), Arr(k + 1, 1)), vbGreen
Next
Call Find_Angle
End Sub

Private Sub Command3_Click()
Form6.Show
End Sub

Private Sub Command4_Click()
Dim FName As String
                On Error GoTo ErrHandler
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Text Files|*.txt|Data Files|*.dat"
CommonDialog1.Flags = cdlOFNOverwritePrompt
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowSave
FName = CommonDialog1.FileName

If CommonDialog1.FileName = "" Then Exit Sub
Call SavePoints(FName, N + 1, Arr())

ErrHandler:
                CommonDialog1.FileName = ""
                Exit Sub

End Sub

Private Sub Command5_Click()
Dim Ans As String
Ans = InputBox("Enter Filename: ", "File Name - ")
If Ans = "" Then Exit Sub
Ans = App.Path + "\Images\" + Ans + ".bmp"
SavePicture Me.Picture1.Image, Ans
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Form_Load()
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
Command5.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button <> 1 Then Exit Sub
If P > N Then Exit Sub
Picture1.PSet (x, y), vbWhite
Picture1.Circle (x, y), 5, vbYellow
Arr(P, 0) = x
Arr(P, 1) = y
Arr(P, 2) = 0
P = P + 1
End Sub
