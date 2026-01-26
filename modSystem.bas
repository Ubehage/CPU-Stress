Attribute VB_Name = "modSystem"
Option Explicit

Private Type SYSTEM_LOGICAL_PROCESSOR_INFORMATION
    ProcessorMask As Long
    Relationship As Long ' 0 = Core, 1 = NUMA, 2 = Cache, 3 = Package
    Reserved(1) As Currency
End Type

Public Type CPU_Info
  PhysicalCores As Long
  KernelsPerCore() As Long
End Type

Private Declare Function GetLogicalProcessorInformation Lib "kernel32" (ByRef Buffer As Any, ByRef ReturnLength As Long) As Long

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef hpvDest As Any, ByRef hpvSource As Any, ByVal cbCopy As Long)

Public Function GetCPUCoreCount() As CPU_Info
  Dim i As Long
  Dim bBuffer() As Byte, cbBuffer As Long, nEntries As Long, structSize As Long, info As SYSTEM_LOGICAL_PROCESSOR_INFORMATION
  Dim bOffset As Long
  Dim r As CPU_Info
  structSize = LenB(info)
  Call GetLogicalProcessorInformation(ByVal 0&, cbBuffer)
  If cbBuffer = 0 Then Exit Function
  ReDim bBuffer(cbBuffer - 1) As Byte
  If GetLogicalProcessorInformation(bBuffer(0), cbBuffer) = 0 Then Exit Function
  nEntries = (cbBuffer / structSize)
  For i = 0 To (nEntries - 1)
    bOffset = (i * structSize)
    CopyMemory info, bBuffer(bOffset), structSize
    Select Case info.Relationship
      Case 0
        If (GetCPUCoreCount.PhysicalCores Mod 5) = 0 Then ReDim Preserve GetCPUCoreCount.KernelsPerCore(1 To (GetCPUCoreCount.PhysicalCores + 5)) As Long
        GetCPUCoreCount.PhysicalCores = (GetCPUCoreCount.PhysicalCores + 1)
        GetCPUCoreCount.KernelsPerCore(GetCPUCoreCount.PhysicalCores) = CountKernelBits(info.ProcessorMask)
    End Select
  Next
  'CopyMemory GetCPUCoreCount, r, LenB(r)
End Function

Private Function CountKernelBits(ByVal kMask As Long) As Long
  Dim i As Long, bMask As Long, r As Long
  bMask = 1
  Do
    If (kMask And bMask) <> 0 Then r = (r + 1)
    If i = 30 Then Exit Do
    bMask = (bMask * 2)
    i = (i + 1)
  Loop
  If (kMask And &H80000000) <> 0 Then r = (r + 1)
  CountKernelBits = r
End Function
