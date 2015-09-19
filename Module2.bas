Attribute VB_Name = "Module2"
Option Explicit
Public N1 As Integer                                'Number of Line segments
Public P1 As Integer                                'Number of Points P = N coz initially P = 0
Public Arr1(0 To MAX_LT, 0 To 2) As Integer

Dim Pos1(0 To MAX_LT, 0 To 2) As Integer                'Used for storing position vectors
Global Theta1(0 To MAX_LT) As Double    'used for storing magnitude of vector
'Function to calculate dot product of two vectors
Private Function dot(ByVal ii As Integer) As Double
Dim i, j As Integer
dot = 0
For i = 0 To N1 - 1
    For j = 0 To 2
    Pos1(i, j) = Arr1(i, j) - Arr1(i + 1, j)
    Next j
Next i
For j = 0 To 2
    dot = dot + Pos1(ii, j) * Pos1(ii + 1, j)
Next j
End Function

'Starting main function
Public Sub Find_Angle1()
Dim i, j, k, L As Integer
Dim Sum As Long
Dim q, x, r As Double

For i = 0 To N1 - 1
Sum = 0
    For j = 0 To 2
    x = ((Arr1(i, j) - Arr1(i + 1, j)) ^ 2)
    Sum = Sum + x
    Next j
Value(i) = Sqr(Sum)
Next i

For i = 0 To N1 - 2
    r = Value(i + 1) * Value(i)
    q = dot(i)
    x = q / r
    If x = 1 Then
        Theta1(i) = 0
    Else
        Theta1(i) = Convert((Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)))
    End If
Next i

End Sub
