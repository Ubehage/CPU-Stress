Attribute VB_Name = "modCPUFixture"
Option Explicit

' Deterministic CPU inventory for mouse/layout tests only. The real CPUView,
' drawing code and timer class are compiled; no stress or CPUID code is linked.
Public Enum CPU_Capabilities
  ccLegacy = 0
  ccSSE2 = 1
  ccAVX = 2
  ccAVX_YMM = 3
End Enum
Public Type CPU_Info
  Name As String
End Type
Global CPUInfo As CPU_Info
Global TotalCores As Long

Public Sub FillCPUInfo()
  TotalCores = 4
  CPUInfo.Name = "Test CPU"
End Sub

Public Function GetCPUCapabilities() As CPU_Capabilities
  GetCPUCapabilities = ccSSE2
End Function
