Attribute VB_Name = "modStress"
Option Explicit

Public Sub StressLoop()
  Dim a As Long, b As Double, c As Long
  Dim Counter As Long
  a = 1
  b = 1.0000001
  Do
    Counter = (Counter + 1)
    c = Sin(b)
    a = (a * c)
    c = Cos(b)
    b = Tan(c)
  Loop Until Counter = 10000000
End Sub
