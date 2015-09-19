Attribute VB_Name = "Module4"
Option Explicit

Public Sub SavePoints(FileName As String, len_arr As Integer, AR As Variant)
Dim i As Integer
Dim FNum As Integer
FNum = FreeFile
Open FileName For Output As #FNum
For i = 0 To len_arr - 1
Print #FNum, AR(i, 0), AR(i, 1), AR(i, 2)
Next i
Print #FNum, len_arr, 0, 0
Close #FNum
End Sub

Public Sub SaveResults(FileName As String, AR As Variant, AR1 As Variant, ByVal MAXscore As Double, l1 As Integer, FileNumber As Integer, l2 As Integer, SCORES As Variant, SCOREL As Variant, FileNME As Variant)
Dim i, j As Integer
Dim FNum As Integer
FNum = FreeFile
FileNumber = FileNumber + 1
Open FileName For Output As #FNum
    Print #FNum, "-----------------Unknown pattern angle values are follows-------------- "
    For i = 0 To l1 - 1
    If AR(i) = Empty Then GoTo Label1
    Print #FNum, AR(i)
    Next i
Label1: Print #FNum, "Number of Angles: ", i
        Print #FNum, "Maximum match score is(Local): ", sco
        Print #FNum, "Maximum match score is(Global): ", MAXscore
        Print #FNum, "Maximum match score(Local) occurred: ", Local_Occur, " times"
        Print #FNum, "--------------Matching for the following angle arrayes--------------"
    For j = 0 To FileNumber
        If FileNME(j) = "" Then GoTo Label3
        Print #FNum, "File is: ", FileNME(j)
        For i = 0 To l2 - 1
        If AR1(j, i) = 0 Or AR1(j, i) = Empty Then GoTo Label2
        Print #FNum, AR1(j, i)
        Next i
Label2: Print #FNum, "Number of Angles: ", i
        Print #FNum, "Match score for Above values is(LOCAL): ", SCORES(j)
        Print #FNum, "Match score for Above values is(GLOBAL): ", SCOREL(j)
    Next j
Label3:
Close #FNum
End Sub

