Attribute VB_Name = "Module7"
Option Explicit
Global Const MAX_LT = 1000
Global Theta1S(0 To MAX_LT, 0 To MAX_LT) As Double
Global Files(0 To MAX_LT) As String
Global MatchScoresL(0 To MAX_LT) As Double
Global MatchScoresS(0 To MAX_LT) As Double
Global del, MAXscore As Double
Global T, Local_Occur As Integer
Global sco As Double
Global Catch(0 To MAX_LT) As Integer

Public Function MAX_(a As Double, ByVal b As Double, Optional ByVal c As Integer) As Double
If a > b Then
MAX_ = a
Else
MAX_ = b
T = c
End If
End Function

Public Sub ClearArr1Theta1()
Dim i As Integer
For i = 0 To MAX_LT
Arr1(i, 0) = Empty
Arr1(i, 1) = Empty
Arr1(i, 2) = Empty
Theta1(i) = Empty
Next i
End Sub

Public Function MAXX_(ARS As Variant, ARL As Variant, ByVal Lt As Integer) As Double ', LS As Integer, LL As Integer) As Double
Dim i, j As Integer
Dim s1 As Double
s1 = 0

For i = 0 To Lt
    sco = MAX_(sco, ARS(i))
Next i

For i = 0 To Lt
    If ARS(i) = sco Then
    s1 = MAX_(s1, ARL(i), i)
    Catch(j) = i
    j = j + 1
    End If
Next i
Local_Occur = j

MAXX_ = s1
End Function
