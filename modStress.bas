Attribute VB_Name = "modStress"
Option Explicit

Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Private Const LOOP_SIZE   As Long = 1000000
Private Const LOOP_COUNT As Long = 100

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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long 'Used to intercept and process VB-window messages (hence the -A variant)
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Public Sub StressLoop()
  Static asmCode() As Byte
  Static hasASM As Boolean
  Static asmAddr As Long
  Dim i As Long
  If hasASM = False Then
    asmCode = SplitAssemblyBytes(GetOptimizedStressAssembly())
    AllowAssemblyExecution asmCode()
    asmAddr = VarPtr(asmCode(0))
    hasASM = True
  End If
  For i = 1 To LOOP_COUNT
    Call CallWindowProc(asmAddr, 0, 0, LOOP_SIZE, 0)
    Call Sleep(0)
  Next
End Sub

Private Function SplitAssemblyBytes(AssemblyBytes As String) As Byte()
  Dim aBytes() As String, b() As Byte, i As Long
  aBytes = Split(AssemblyBytes, " ")
  ReDim b(0 To UBound(aBytes)) As Byte
  For i = 0 To UBound(aBytes)
    b(i) = CByte("&H" & aBytes(i))
  Next
  SplitAssemblyBytes = b
End Function

Private Function GetOptimizedStressAssembly() As String
  Select Case GetCPUCapabilities()
    Case CPU_Capabilities.ccAVX
      GetOptimizedStressAssembly = "55 89 E5 53 57 8B 7D 10 B8 01 00 00 00 F2 0F 2A C8 0F 28 C1 0F 28 D1 0F 28 D9 0F 28 E1 0F 28 E9 0F 28 F1 0F 28 F9 0F 59 C1 0F 59 D1 0F 59 D9 0F 59 E1 0F 59 E9 0F 59 F1 0F 59 F9 0F 58 C1 0F 58 D1 0F 58 D9 0F 58 E1 0F 58 E9 0F 58 F1 0F 58 F9 4F 75 D1 5F 5B 5D C2 10 00"
      'push ebp
      'mov ebp, esp
      'push ebx
      'push edi
      'mov edi, [ebp+16]
      'mox eax, 1
      'cvtsi2sd xmm1, eax
      'movaps xmm0, xmm1
      'movaps xmm2, xmm1
      'movaps xmm3, xmm1
      'movaps xmm4, xmm1
      'movaps xmm5, xmm1
      'movaps xmm6, xmm1
      'movaps xmm7, xmm1
      'mulps xmm0, xmm1
      'mulps xmm2, xmm1
      'mulps xmm3, xmm1
      'mulps xmm4, xmm1
      'mulps xmm5, xmm1
      'mulps xmm6, xmm1
      'mulps xmm7, xmm1
      'addps xmm0, xmm1
      'addps xmm2, xmm1
      'addps xmm3, xmm1
      'addps xmm4, xmm1
      'addps xmm5, xmm1
      'addps xmm6, xmm1
      'addps xmm7, xmm1
      'dec edi
      'jnz -47
      'pop edi
      'pop ebx
      'pop ebp
      'ret 16
    Case CPU_Capabilities.ccSSE2
      GetOptimizedStressAssembly = "55 89 E5 57 8B 7D 10 B8 01 00 00 00 F2 0F 2A C8 0F 28 C1 0F 59 C0 0F 58 C1 0F 59 C0 0F 58 C1 4F 75 F1 5F 5D C2 10 00"
      'push ebp
      'mov ebp, esp
      'push ebx
      'push edi
      'mov edi, [ebp+16]
      'mox eax, 1
      'cvtsi2sd xmm1, eax
      'movaps xmm0, xmm1
      'mulps xmm0, xmm0
      'addps xmm0, xmm1
      'mulps xmm0, xmm0
      'addps xmm0, xmm1
      'dec edi
      'jnz -15
      'pop edi
      'pop ebp
      'ret 16
    Case CPU_Capabilities.ccLegacy
      GetOptimizedStressAssembly = "55 89 E5 53 57 8B 7D 10 D9 E8 D9 E8 DE F1 DE F1 DE F1 DD D8 4F 75 F1 5F 5B 5D C2 10 00"
      'push ebp
      'mov ebp, esp
      'push ebx
      'push edi
      'mov edi, [ebp+16]
      'fld1
      'fld1
      'fdivp st(1), st(0)
      'fdivp st(1), st(0)
      'fdivp st(1), st(0)
      'fstp st(0)
      'dec edi
      'jnz -15
      'pop edi
      'pop ebp
      'ret 16
  End Select
End Function

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

Private Sub AllowAssemblyExecution(AssemblyArray() As Byte)
  Dim oProtect As Long
  Call VirtualProtect(AssemblyArray(0), UBound(AssemblyArray) + 1, PAGE_EXECUTE_READWRITE, oProtect)
End Sub
