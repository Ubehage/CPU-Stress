Attribute VB_Name = "modStress"
Option Explicit

Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Private Const LOOP_SIZE   As Long = 1000000
Private Const LOOP_COUNT As Long = 100

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long 'Used to intercept and process VB-window messages (hence the -A variant)
Private Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long

Public Sub StressLoop()
  Static asmCode() As Byte
  Static hasASM As Boolean
  Dim i As Long
  If hasASM = False Then
    Dim oProtect As Long
    asmCode = SplitAssemblyBytes("55 89 E5 57 8B 7D 10 B8 01 00 00 00 F2 0F 2A C8 0F 28 C1 0F 59 C0 0F 58 C1 0F 59 C0 0F 58 C1 4F 75 F1 5F 5D C2 10 00")
    Call VirtualProtect(asmCode(0), UBound(asmCode) + 1, PAGE_EXECUTE_READWRITE, oProtect)
    hasASM = True
  End If
  For i = 1 To LOOP_COUNT
    Call CallWindowProc(VarPtr(asmCode(0)), 0, 0, LOOP_SIZE, 0)
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
