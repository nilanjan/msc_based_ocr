VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Modified Needleman-Wunsch Pattern Matching Technique"
   ClientHeight    =   2880
   ClientLeft      =   5265
   ClientTop       =   4845
   ClientWidth     =   5670
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5670
   Begin VB.CommandButton Command1 
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
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.DriveListBox Drive2 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   2535
   End
   Begin VB.DirListBox Dir2 
      Height          =   1665
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.FileListBox File2 
      Height          =   1650
      Left            =   3000
      Pattern         =   "*.txt;*.dat"
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select the DB Folder"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim Ans As String

Ans = InputBox("Enter value of delta(in Degree): ", "DELTA")
If Ans = "" Then Exit Sub
del = Val(Ans)


Dim i, j, L As Integer
L = File2.ListCount

For i = 0 To L - 1
Files(i) = File2.Path + "\" + File2.List(i)
Next i

For i = 0 To L - 1
N1 = OpenFiles(Files(i), Arr1())
Call Find_Angle1
        For j = 0 To MAX_LT
        Theta1S(i, j) = Theta1(j)
        Next j
MatchScoresS(i) = N_Match_S(del)
MatchScoresL(i) = N_Match_L(del)
Call ClearArr1Theta1
Next i

MAXscore = MAXX_(MatchScoresS(), MatchScoresL(), L - 1)

Form2.Label2.Caption = "Match (Global) is: " + Str$(Int(MAXscore)) + "%"
Form2.Label3.Caption = "Match (Local) is: " + Str$(Int(sco)) + "%"
Form2.Label1.Caption = "Best Match Pattern File is: " + Files(T)
Form2.Label4.Caption = "Local Match Occurrance is: " + Str$(Local_Occur)

If Local_Occur = 1 Then
Form2.Command1.Enabled = False
Else
Form2.Command1.Enabled = True
End If

'Codes for Displaying The Best Match pattern.
N1 = OpenFiles(Files(T), Arr1())
For i = 0 To N1
Call SetPoints(Arr1(i, 0), Arr1(i, 1), Form2.Picture1)
Next i
For i = 0 To N1 - 1
    Form2.Picture1.Line (Arr1(i, 0), Arr1(i, 1))-(Arr1(i + 1, 0), Arr1(i + 1, 1)), vbGreen
Next

Form2.Show
Me.Hide
End Sub

Private Sub Form_Load()
Drive2.Drive = App.Path
Dir2.Path = App.Path

End Sub

Private Sub Dir2_Change()
On Error GoTo ErrorHandler
File2.Path = Dir2.Path

Exit Sub
ErrorHandler:
MsgBox Err.Description, vbOKOnly, "Error!"
End Sub

Private Sub Drive2_Change()
On Error GoTo ErrHandler
Dir2.Path = Drive2.Drive

Exit Sub
ErrHandler:
MsgBox Err.Description, vbOKOnly, "Error!"
End Sub
