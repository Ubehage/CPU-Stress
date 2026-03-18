Attribute VB_Name = "modStress"
Option Explicit

Private Const LOOP_SIZE   As Long = 100000
Private Const LOOP_COUNT As Long = 100

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub StressLoop()
  Dim a As Double, b As Double, c As Double
  Dim v1 As Double, v2 As Double
  Dim Counter As Long, i As Long
  v1 = 1.23456789
  v2 = 0.0000001
  For i = 1 To LOOP_COUNT
    Counter = 0
    Do
      a = Log(v1 + a)
      b = (Sin(a) * Cos(a + 0.1))
      c = Exp(b)
      a = Sqr(Abs(c))
      a = (a / 0.999999)
      If a > 1000000 Then a = 0.1
      a = (a + (a * v2))
      Counter = Counter + 1
    Loop Until Counter >= LOOP_SIZE
    Call Sleep(0)
  Next
End Sub
