Attribute VB_Name = "modLabWOW64"
Option Explicit

Private Declare Function IsWow64Process Lib "kernel32" (ByVal Process As Long, ByRef IsWow64 As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Sub GetNativeSystemInfo Lib "kernel32" (ByRef SystemInfo As SYSTEM_INFO)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type SYSTEM_INFO
    Architecture As Integer
    Reserved As Integer
    PageSize As Long
    MinimumAddress As Long
    MaximumAddress As Long
    ProcessorMask As Long
    ProcessorCount As Long
    ProcessorType As Long
    AllocationGranularity As Long
    ProcessorLevel As Integer
    ProcessorRevision As Integer
End Type

' This runs in a disposable 32-bit worker, never in the GUI process.
' The code uses the Windows x64 ABI but may not call 64-bit Windows APIs.
Public Function BuildWOW64Code(ByRef Payload() As Byte) As Byte()
    Dim Wow64 As Long, Info As SYSTEM_INFO
    Dim FileNumber As Integer, GateSize As Long, PayloadSize As Long
    Dim Combined() As Byte, Gate() As Byte, Detail As String
    On Error GoTo Failed
    If IsWow64Process(GetCurrentProcess(), Wow64) = 0 Then Err.Raise vbObjectError + 2300, , "Cannot query WOW64 process state."
    GetNativeSystemInfo Info
    If Wow64 = 0 Or Info.Architecture <> 9 Then Err.Raise vbObjectError + 2301, , "Direct x64 execution requires a 32-bit process on AMD64 Windows (WOW64)."
    FileNumber = FreeFile
    Open App.Path & "\wow64-gate.bin" For Binary Access Read As #FileNumber
    GateSize = LOF(FileNumber)
    If GateSize < 1 Or GateSize > 4096 Then Err.Raise vbObjectError + 2302, , "Invalid or missing WOW64 bridge binary. Rebuild the lab."
    ReDim Gate(0 To GateSize - 1)
    Get #FileNumber, , Gate
    Close #FileNumber
    FileNumber = 0
    PayloadSize = UBound(Payload) - LBound(Payload) + 1
    ReDim Combined(0 To GateSize + PayloadSize - 1)
    CopyMemory Combined(0), Gate(0), GateSize
    CopyMemory Combined(GateSize), Payload(LBound(Payload)), PayloadSize
    BuildWOW64Code = Combined
    Exit Function
Failed:
    Detail = Err.Description
    On Error Resume Next
    If FileNumber <> 0 Then Close #FileNumber
    On Error GoTo 0
    Err.Raise vbObjectError + 2303, "BuildWOW64Code", Detail
End Function
