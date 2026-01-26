Attribute VB_Name = "modMain"
Option Explicit

Sub Main()
  Dim i As Long, c As Long
  With GetCPUCoreCount
    Debug.Print "CPU Cores: " & CStr(.PhysicalCores)
    For i = 1 To .PhysicalCores
      Debug.Print "Core(" & CStr(i) & "): " & CStr(.KernelsPerCore(i)) & " kernels."
      c = (c + .KernelsPerCore(i))
    Next
    Debug.Print "Total kernels: " & CStr(c)
  End With
End Sub
