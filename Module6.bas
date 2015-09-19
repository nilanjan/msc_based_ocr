Attribute VB_Name = "Module6"
Option Explicit
Public Sub SetPoints(ByVal x As Integer, ByVal y As Integer, F As Control)
F.PSet (x, y), vbWhite
F.Circle (x, y), 5, vbYellow
End Sub
