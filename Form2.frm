VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   Caption         =   "Modified Needleman-Wunsch Pattern Matching Technique - Shape #2"
   ClientHeight    =   6945
   ClientLeft      =   3915
   ClientTop       =   1020
   ClientWidth     =   7785
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   7785
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Best Match"
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
      TabIndex        =   8
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Other Matches"
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
      Top             =   6360
      Width           =   1335
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
      Height          =   615
      Left            =   6360
      MaskColor       =   &H0000FF00&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2880
      Top             =   3240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
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
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
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
      TabIndex        =   1
      Top             =   120
      Width           =   1335
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
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
      Height          =   975
      Left            =   6360
      TabIndex        =   9
      Top             =   4560
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   975
      Left            =   6360
      TabIndex        =   6
      Top             =   2160
      Width           =   1335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   975
      Left            =   6360
      TabIndex        =   3
      Top             =   3360
      Width           =   1335
      WordWrap        =   -1  'True
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
      TabIndex        =   2
      Top             =   120
      Width           =   6135
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

'Codes for Displaying The Best Match pattern.
For j = 0 To Local_Occur - 1
Me.Picture1.Cls
N1 = OpenFiles(Files(Catch(j)), Arr1())

Form2.Label2.Caption = "Match (Global) is: " + Str$(Int(MatchScoresL(Catch(j)))) + "%"
Form2.Label3.Caption = "Match (Local) is: " + Str$(Int(MatchScoresS(Catch(j)))) + "%"
Form2.Label1.Caption = "Pattern File is: " + Files(Catch(j))

For i = 0 To N1
Call SetPoints(Arr1(i, 0), Arr1(i, 1), Form2.Picture1)
Next i
For i = 0 To N1 - 1
    Form2.Picture1.Line (Arr1(i, 0), Arr1(i, 1))-(Arr1(i + 1, 0), Arr1(i + 1, 1)), vbGreen
Next i
MsgBox "To see next Match Click Ok", vbInformation, "Match No." & j
Next j

End Sub

Private Sub Command2_Click()
Me.Picture1.Cls
Form2.Label2.Caption = "Match (Global) is: " + Str$(Int(MAXscore)) + "%"
Form2.Label3.Caption = "Match (Local) is: " + Str$(Int(sco)) + "%"
Form2.Label1.Caption = "Best Match Pattern File is: " + Files(T)

'Codes for Displaying The Best Match pattern.
N1 = OpenFiles(Files(T), Arr1())
For i = 0 To N1
Call SetPoints(Arr1(i, 0), Arr1(i, 1), Form2.Picture1)
Next i
For i = 0 To N1 - 1
    Form2.Picture1.Line (Arr1(i, 0), Arr1(i, 1))-(Arr1(i + 1, 0), Arr1(i + 1, 1)), vbGreen
Next

End Sub

Private Sub Command3_Click()
Dim FName As String
            On Error GoTo ErrHandler
CommonDialog1.FileName = ""
CommonDialog1.Filter = "Text Files|*.txt|Data Files|*.dat"
CommonDialog1.Flags = cdlOFNOverwritePrompt
CommonDialog1.InitDir = App.Path
CommonDialog1.ShowSave
FName = CommonDialog1.FileName

If CommonDialog1.FileName = "" Then Exit Sub
Call SaveResults(FName, Theta(), Theta1S(), MAXscore, UBound(Theta()), UBound(Theta1S(), 1), UBound(Theta1S(), 2), MatchScoresS(), MatchScoresL(), Files())

ErrHandler:
                CommonDialog1.FileName = ""
                Exit Sub
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
Dim Ans As String
Ans = InputBox("Enter Filename: ", "File Name - ")
If Ans = "" Then Exit Sub
Ans = App.Path + "\Images\" + Ans + ".bmp"
SavePicture Me.Picture1.Image, Ans
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
