Attribute VB_Name = "modStress"
Option Explicit

Private Const LOOP_COUNTER As Long = 5000000

Public Sub StressLoop()
  Dim a As Double, b As Double, c As Double
  Dim v1 As Double, v2 As Double
  Dim Counter As Long
  v1 = 1.23456789
  v2 = 0.0000001
  Do
    a = Log(v1 + a)
    b = (Sin(a) * Cos(a + 0.1))
    c = Exp(b)
    a = Sqr(Abs(c))
    a = (a / 0.999999)
    If a > 1000000 Then a = 0.1
    a = (a + (a * v2))
    Counter = (Counter + 1)
  Loop Until Counter >= LOOP_COUNTER
End Sub
