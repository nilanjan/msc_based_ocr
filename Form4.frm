VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   Caption         =   "Input Type"
   ClientHeight    =   1455
   ClientLeft      =   6405
   ClientTop       =   1920
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4680
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Input from File."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Input using Mouse."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Integer
Dim Ans, FName As String

On Error GoTo ErrHandler

If Option1.Value = True Then
        Ans = InputBox("Enter number of lines: ", "Line Numbers - ")
    If Ans = "" Then Exit Sub
    N = Val(Ans)
        Form1.Picture1.Enabled = True
        Form1.Label1.Caption = "Input taken using Mouse."
ElseIf Option2.Value = True Then
    
    Form1.Command4.Enabled = False
    
        CommonDialog1.InitDir = App.Path
        CommonDialog1.FileName = ""
        CommonDialog1.Filter = "Text Files|*.txt|Data Files|*.dat"
        CommonDialog1.Flags = cdlOFNFileMustExist
        CommonDialog1.ShowOpen
        If CommonDialog1.FileName = "" Then Exit Sub
        FName = CommonDialog1.FileName
        Form1.Label1.Caption = "Unknown Pattern File is: " + FName
        N = OpenFiles(FName, Arr())
            For i = 0 To N
            Call SetPoints(Arr(i, 0), Arr(i, 1), Form1.Picture1)
            Next i
End If
Me.Hide
Form1.Enabled = True

ErrHandler:
CommonDialog1.FileName = ""
End Sub
