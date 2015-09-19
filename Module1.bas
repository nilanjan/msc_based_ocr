Attribute VB_Name = "Module1"
Option Explicit
Public N As Integer                                'Number of Line segments
Public P As Integer                                'Number of Points P = N coz initially P = 0
Public Arr(0 To MAX_LT, 0 To 2) As Integer            'Co-ordinates
Const PI = 22 / 7

Dim Pos(0 To MAX_LT, 0 To 2) As Integer                'Used for storing position vectors
Global Theta(0 To MAX_LT), Value(0 To MAX_LT) As Double     'used for storing magnitude of vector
        
'Function to convert radian to degree
Public Function Convert(ByVal T As Double) As Double
Convert = (180 * T) / PI
End Function

'Function to calculate dot product of two vectors
Private Function dot(ByVal ii As Integer) As Double
Dim i, j As Integer
dot = 0
For i = 0 To N - 1
    For j = 0 To 2
    Pos(i, j) = Arr(i, j) - Arr(i + 1, j)
    Next j
Next i
For j = 0 To 2
    dot = dot + Pos(ii, j) * Pos(ii + 1, j)
Next j
End Function

'Starting main function
Public Sub Find_Angle()
Dim i, j, k, L As Integer
Dim Sum As Long
Dim q, x, r As Double

For i = 0 To N - 1
Sum = 0
    For j = 0 To 2
    x = ((Arr(i, j) - Arr(i + 1, j)) ^ 2)
    Sum = Sum + x
    Next j
Value(i) = Sqr(Sum)
Next i

For i = 0 To N - 2
    r = Value(i + 1) * Value(i)
    q = dot(i)
    x = q / r
    If x = 1 Then
        Theta(i) = 0
    Else
        Theta(i) = Convert((Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)))
    End If
Next i

End Sub
