Attribute VB_Name = "modCPU"
Option Explicit

Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Public Enum CPU_Capabilities
  ccAVX = &H1
  ccSSE2 = &H2
  ccLegacy = &H3
End Enum

Private Type CPUID_Result
  dwEAX As Long
  dwEBX As Long
  dwECX As Long
  dwEDX As Long
End Type

Public Type CPU_Info
  Name As String
  PhysicalCores As Long
  KernelsPerCore() As Long
End Type

Private Type SYSTEM_LOGICAL_PROCESSOR_INFORMATION
  ProcessorMask As Long
  Relationship As Long ' 0 = Core, 1 = NUMA, 2 = Cache, 3 = Package
  Reserved(1) As Currency
End Type

Private Declare Function SetProcessAffinityMask Lib "kernel32" (ByVal hProcess As Long, ByVal dwProcessAffinityMask As Long) As Long

Private Declare Function GetLogicalProcessorInformation Lib "kernel32" (ByRef Buffer As Any, ByRef ReturnLength As Long) As Long

Global CPUInfo As CPU_Info
Global TotalCores As Long

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
End Function

Public Function CountAllCores(cInfo As CPU_Info) As Long
  Dim i As Long, r As Long
  For i = 1 To cInfo.PhysicalCores
    r = (r + cInfo.KernelsPerCore(i))
  Next
  CountAllCores = r
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

Public Function GetCPUName() As String
  Dim hKey As Long, res As Long, sPath As String, sValue As String, nSize As Long
  sPath = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"
  res = RegOpenKeyEx(HKEY_LOCAL_MACHINE, sPath, 0, KEY_READ, hKey)
  If res = 0 Then
    Call RegQueryValueEx(hKey, "ProcessorNameString", 0, 0, ByVal 0&, nSize)
    If nSize > 0 Then
      sValue = String$(nSize, Chr$(0))
      Call RegQueryValueEx(hKey, "ProcessorNameString", 0, 0, ByVal sValue, nSize)
      GetCPUName = Trim$(Left$(sValue, (InStr(sValue, Chr$(0)) - 1)))
    Else
      GetCPUName = "Unknown Processor"
    End If
    Call RegCloseKey(hKey)
  Else
    GetCPUName = "Unknown Processor"
  End If
End Function

Public Sub LockProcessToCPU(CPUIndex As Long)
  If (CPUIndex < 0 Or CPUIndex > (TotalCores - 1)) Then Exit Sub
  Dim hProc As Long, cMask As Long
  hProc = GetCurrentProcess()
  If CPUIndex <= 30 Then
    cMask = 2 ^ CPUIndex
  ElseIf CPUIndex = 31 Then
    cMask = &H80000000
  Else
    'Sorry, VB6 cannot go any higher
  End If
  Call SetProcessAffinityMask(hProc, cMask)
End Sub

Public Sub FillCPUInfo()
  CPUInfo = GetCPUCoreCount()
  CPUInfo.Name = GetCPUName()
  TotalCores = CountAllCores(CPUInfo)
End Sub

Public Function GetCPUCapabilities() As CPU_Capabilities
  Dim cpuCode() As Byte
  Dim CPUResult As CPUID_Result
  cpuCode = SplitAssemblyBytes("55 89 E5 53 57 8B 45 08 8B 7D 0C 31 C9 0F A2 89 07 89 5F 04 89 4F 08 89 57 0C 5F 5B 5D C2 10 00")
  AllowAssemblyExecution cpuCode()
  Call CallWindowProc(VarPtr(cpuCode(0)), 1, VarPtr(CPUResult), 0, 0)
  If (CPUResult.dwECX And &H10000000) Then
    GetCPUCapabilities = ccAVX
  ElseIf (CPUResult.dwEDX And &H4000000) Then
    GetCPUCapabilities = ccSSE2
  Else
    GetCPUCapabilities = ccLegacy
  End If
End Function

Public Sub AllowAssemblyExecution(AssemblyArray() As Byte)
  Dim unused As Long
  Call VirtualProtect(AssemblyArray(0), UBound(AssemblyArray) + 1, PAGE_EXECUTE_READWRITE, unused)
End Sub

Public Function SplitAssemblyBytes(AssemblyBytes As String) As Byte()
  Dim aBytes() As String, b() As Byte, i As Long
  aBytes = Split(AssemblyBytes, " ")
  If UBound(aBytes) = 0 Then Exit Function
  ReDim b(0 To UBound(aBytes)) As Byte
  For i = 0 To UBound(aBytes)
    b(i) = CByte("&H" & aBytes(i))
  Next
  SplitAssemblyBytes = b
End Function
