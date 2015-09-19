Attribute VB_Name = "Module3"
Const MAX_STR_LEN = 255
Const W = 0
Const Match = 1
Const MisMatch = 0

Function S1(ByVal a As Double, ByVal b As Double, ByVal DELTA As Double) As Double
    If Abs(a - b) < DELTA Then
    S1 = Match '1 / (1 + Abs(a - b) * Abs(a - b))
    Else
    S1 = MisMatch '1 / (1 + Abs(a - b) * Abs(a - b))
    End If
End Function

Function MAX(ByVal a As Double, ByVal b As Double, ByVal c As Double) As Double
    If (a > b) Then
    MAX = a
    Else
    MAX = b
    End If
    If (c > MAX) Then
    MAX = c
    End If
End Function

Function N_Match_L(ByVal del As Double) As Double
Dim temp As Double
Dim i, j As Integer
N = N - 1
N1 = N1 - 1
ReDim M(0 To N, 0 To N1) As Double
'Initialization Step------------------------------------------------------------------
        For j = 0 To N1
        M(0, j) = 0
        Next j
        For j = 0 To N
        M(j, 0) = 0
        Next j
'Matrix Fill-Up Step-------------------------------------------------------------
    For i = 1 To N
        For j = 1 To N1
        temp = S1(Theta(i - 1), Theta1(j - 1), del)
        M(i, j) = MAX(M(i - 1, j - 1) + temp, M(i, j - 1) + W, M(i - 1, j) + W)
        Next j
    Next i

N_Match_L = M(N, N1)
    If N > N1 Then
    N_Match_L = (N_Match_L / (N * Match)) * 100
    Else
    N_Match_L = (N_Match_L / (N1 * Match)) * 100
    End If
N = N + 1
End Function

Function N_Match_S(ByVal del As Double) As Double
Dim temp As Double
Dim i, j As Integer
N = N - 1
N1 = N1 - 1
ReDim M(0 To N, 0 To N1) As Double
'Initialization Step------------------------------------------------------------------
        For j = 0 To N1
        M(0, j) = 0
        Next j
        For j = 0 To N
        M(j, 0) = 0
        Next j
'Matrix Fill-Up Step-------------------------------------------------------------
    For i = 1 To N
        For j = 1 To N1
        temp = S1(Theta(i - 1), Theta1(j - 1), del)
        M(i, j) = MAX(M(i - 1, j - 1) + temp, M(i, j - 1) + W, M(i - 1, j) + W)
        Next j
    Next i
N_Match_S = M(N, N1)
    If N < N1 Then
    N_Match_S = (N_Match_S / (N * Match)) * 100
    Else
    N_Match_S = (N_Match_S / (N1 * Match)) * 100
    End If
N = N + 1
End Function

