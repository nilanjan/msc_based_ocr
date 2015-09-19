Attribute VB_Name = "Module5"
Option Explicit
Public Function OpenFiles(FileName As String, AR As Variant) As Integer
Dim i As Integer
Dim FNum As Integer
FNum = FreeFile
Open FileName For Input As #FNum
i = 0
Do Until EOF(FNum)
    Input #FNum, AR(i, 0), AR(i, 1), AR(i, 2)
    i = i + 1
Loop
OpenFiles = AR(i - 1, 0) - 1
AR(i - 1, 0) = 0
Close #FNum
End Function
