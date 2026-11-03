Attribute VB_Name = "modStress"
Option Explicit

' Compiled stress-worker execution only. Each call returns after a calibrated
' batch so the worker can process stop/restart/exit commands between batches.
Private Declare Function StressVirtualAlloc Lib "kernel32" Alias "VirtualAlloc" (ByVal Address As Long, ByVal Size As Long, ByVal AllocationType As Long, ByVal Protection As Long) As Long
Private Declare Function StressVirtualProtect Lib "kernel32" Alias "VirtualProtect" (ByVal Address As Long, ByVal Size As Long, ByVal Protection As Long, ByRef Previous As Long) As Long
Private Declare Function StressVirtualFree Lib "kernel32" Alias "VirtualFree" (ByVal Address As Long, ByVal Size As Long, ByVal FreeType As Long) As Long
Private Declare Function StressFlushInstructionCache Lib "kernel32" Alias "FlushInstructionCache" (ByVal Process As Long, ByVal Address As Long, ByVal Size As Long) As Long
Private Declare Function StressGetCurrentProcess Lib "kernel32" Alias "GetCurrentProcess" () As Long
Private Declare Function StressIsWow64Process Lib "kernel32" Alias "IsWow64Process" (ByVal Process As Long, ByRef IsWow64 As Long) As Long
Private Declare Sub StressGetNativeSystemInfo Lib "kernel32" Alias "GetNativeSystemInfo" (Information As Any)
Private Declare Function StressIsProcessorFeaturePresent Lib "kernel32" Alias "IsProcessorFeaturePresent" (ByVal Feature As Long) As Long
Private Declare Function StressCallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal Address As Long, ByVal Arg1 As Long, ByVal Arg2 As Long, ByVal Arg3 As Long, ByVal Arg4 As Long) As Long
Private Declare Sub StressCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Size As Long)
Private Declare Function StressQueryPerformanceFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef Frequency As Currency) As Long
Private Declare Function StressQueryPerformanceCounter Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef Counter As Currency) As Long

Private Const MEM_COMMIT_RESERVE As Long = &H3000&
Private Const MEM_RELEASE As Long = &H8000&
Private Const PAGE_READWRITE As Long = &H4&
Private Const PAGE_EXECUTE_READ As Long = &H20&
Private Const SCRATCH_SIZE As Long = 4096
Private Const INITIAL_BATCH As Long = 4096
Private Const MAX_BATCH As Long = 8000000
Private Const TARGET_MS As Double = 15#

Private m_Code As Long
Private m_Scratch As Long
Private m_Initialized As Boolean
Private m_Is64Bit As Boolean
Private m_KernelName As String
Private m_BatchIterations As Long
Private m_Frequency As Currency
Private m_LastElapsedMs As Double
Private m_ReturnLow As Long
Private m_ReturnHigh As Long

Public Sub StressLoop()
    EnsureStressInitialized
    ExecuteStressBatch m_BatchIterations
    CalibrateNextBatch
End Sub

' Exact-count integration test entry point; leaves batch calibration unchanged.
Public Function StressRunBatch(ByVal Iterations As Long) As String
    If Iterations < 1 Or Iterations > MAX_BATCH Then Err.Raise 5, "StressRunBatch", "Iterations must be between 1 and 8000000."
    EnsureStressInitialized
    ExecuteStressBatch Iterations
    If m_Is64Bit Then StressRunBatch = Right$("00000000" & Hex$(m_ReturnHigh), 8)
    StressRunBatch = StressRunBatch & Right$("00000000" & Hex$(m_ReturnLow), 8)
End Function

Public Function StressKernelName() As String
    StressKernelName = m_KernelName
    If Not m_Initialized Then StressKernelName = "Not initialized"
End Function

' Number of iterations planned for the next StressLoop call.
Public Function StressBatchIterations() As Long
    StressBatchIterations = m_BatchIterations
    If StressBatchIterations = 0 Then StressBatchIterations = INITIAL_BATCH
End Function

Public Function StressLastElapsedMs() As Double
    StressLastElapsedMs = m_LastElapsedMs
End Function

' First 32 scratch bytes in memory order, useful for finite-state validation.
Public Function StressReadScratch() As String
    Dim Bytes(0 To 31) As Byte, i As Long
    If m_Scratch = 0 Then Exit Function
    StressCopyMemory Bytes(0), ByVal m_Scratch, 32
    For i = 0 To 31
        StressReadScratch = StressReadScratch & Right$("0" & Hex$(Bytes(i)), 2)
    Next i
End Function

' The bridge records its native execution selector independently of the kernel.
Public Function StressCodeSegment() As Long
    Dim Selector As Integer
    If Not m_Is64Bit Or m_Scratch = 0 Then Exit Function
    StressCopyMemory Selector, ByVal (m_Scratch + 4088), 2
    StressCodeSegment = CLng(Selector) And &HFFFF&
End Function

Public Sub StressShutdown()
    If m_Code <> 0 Then Call StressVirtualFree(m_Code, 0, MEM_RELEASE)
    If m_Scratch <> 0 Then Call StressVirtualFree(m_Scratch, 0, MEM_RELEASE)
    m_Code = 0
    m_Scratch = 0
    m_Initialized = False
    m_Is64Bit = False
    m_KernelName = vbNullString
    m_BatchIterations = 0
    m_LastElapsedMs = 0
    m_ReturnLow = 0
    m_ReturnHigh = 0
    m_Frequency = 0
End Sub

Private Sub EnsureStressInitialized()
    Dim Registers(0 To 3) As Long, NativeInfo(0 To 8) As Long
    Dim XCR0(0 To 1) As Long, IsWow64 As Long
    Dim CanSSE2 As Boolean, CanAVX As Boolean, CanFMA As Boolean
    Dim Probe As Long, Payload As String, Detail As String
    On Error GoTo Failed
    If IsRunningInIDE Then Err.Raise 5, "StressLoop", "Machine-code stress is disabled inside the VB6 IDE. Compile and run the application."
    If m_Initialized Then Exit Sub
    If StressQueryPerformanceFrequency(m_Frequency) = 0 Or m_Frequency <= 0 Then Err.Raise 5, , "Cannot initialize stress timing."
    Probe = AllocateExecutable(CPUIDProbeHex())
    Call StressCallWindowProc(Probe, 0, VarPtr(Registers(0)), 0, 0)
    If Registers(0) >= 1 Then
        Call StressCallWindowProc(Probe, 1, VarPtr(Registers(0)), 0, 0)
        CanSSE2 = ((Registers(3) And &H4000000) <> 0)
        CanSSE2 = CanSSE2 And (StressIsProcessorFeaturePresent(10) <> 0)
        CanAVX = ((Registers(2) And &H1C000000) = &H1C000000)
        CanFMA = ((Registers(2) And &H1000&) <> 0)
    End If
    Call StressVirtualFree(Probe, 0, MEM_RELEASE)
    Probe = 0
    ' XGETBV is legal only after checking XSAVE and OSXSAVE; require AVX too.
    If CanAVX Then
        Probe = AllocateExecutable(XGETBVProbeHex())
        Call StressCallWindowProc(Probe, 0, VarPtr(XCR0(0)), 0, 0)
        CanAVX = ((XCR0(0) And &H6&) = &H6&)
        Call StressVirtualFree(Probe, 0, MEM_RELEASE)
        Probe = 0
    End If
    CanFMA = CanFMA And CanAVX
    StressGetNativeSystemInfo NativeInfo(0)
    If StressIsWow64Process(StressGetCurrentProcess(), IsWow64) = 0 Then Err.Raise 5, , "Cannot identify stress worker architecture."
    ' Selector bridge only on x86 WOW64 hosted by native AMD64 Windows.
    ' ARM64 and native 32-bit Windows use an x86 kernel instead.
    m_Is64Bit = IsWow64 <> 0 And ((NativeInfo(0) And &HFFFF&) = 9) And CanSSE2
    If m_Is64Bit Then
        If CanFMA Then
            Payload = StressFMA64Hex()
            m_KernelName = "WOW64 x64 FMA"
        ElseIf CanAVX Then
            Payload = StressAVX64Hex()
            m_KernelName = "WOW64 x64 AVX"
        Else
            Payload = StressSSE64Hex()
            m_KernelName = "WOW64 x64 SSE2"
        End If
        Payload = StressWOW64GateHex() & " " & Payload
    ElseIf CanSSE2 Then
        Payload = StressSSE32Hex()
        m_KernelName = "x86 SSE2"
    Else
        Payload = StressLegacy32Hex()
        m_KernelName = "x86 legacy"
    End If
    m_Code = AllocateExecutable(Payload)
    m_Scratch = StressVirtualAlloc(0, SCRATCH_SIZE, MEM_COMMIT_RESERVE, PAGE_READWRITE)
    If m_Scratch = 0 Then Err.Raise 7, , "Cannot allocate stress scratch memory."
    m_BatchIterations = INITIAL_BATCH
    m_Initialized = True
    Exit Sub
Failed:
    Detail = Err.Description
    If Probe <> 0 Then Call StressVirtualFree(Probe, 0, MEM_RELEASE)
    StressShutdown
    Err.Raise vbObjectError + 2300, "StressLoop", Detail
End Sub

Private Sub ExecuteStressBatch(ByVal Iterations As Long)
    Dim Started As Currency, Finished As Currency
    If Iterations < 1 Or Iterations > MAX_BATCH Then Err.Raise 5, , "Invalid stress batch size."
    If StressQueryPerformanceCounter(Started) = 0 Then Err.Raise 5, , "Cannot start stress timing."
    m_ReturnLow = StressCallWindowProc(m_Code, 0, m_Scratch, Iterations, SCRATCH_SIZE)
    If StressQueryPerformanceCounter(Finished) = 0 Then Err.Raise 5, , "Cannot finish stress timing."
    m_LastElapsedMs = (CDbl(Finished) - CDbl(Started)) * 1000# / CDbl(m_Frequency)
    If m_Is64Bit Then
        StressCopyMemory m_ReturnLow, ByVal (m_Scratch + 4080), 4
        StressCopyMemory m_ReturnHigh, ByVal (m_Scratch + 4084), 4
    Else
        m_ReturnHigh = 0
    End If
End Sub

Private Sub CalibrateNextBatch()
    Dim NextCount As Double, GrowthLimit As Double
    ' Target 15 ms; reduce more aggressively after an observed >25 ms batch.
    ' Scheduling/clock changes can extend a synchronous call, so this measured
    ' target is not a hard real-time interrupt of machine code.
    If m_LastElapsedMs <= 0 Then Exit Sub
    NextCount = CDbl(m_BatchIterations) * TARGET_MS / m_LastElapsedMs
    If m_LastElapsedMs > 25# Then NextCount = NextCount * 0.8
    GrowthLimit = CDbl(m_BatchIterations) * 2#
    If m_LastElapsedMs >= 8# Then GrowthLimit = CDbl(m_BatchIterations) * 1.25
    If NextCount > GrowthLimit Then NextCount = GrowthLimit
    If NextCount > MAX_BATCH Then NextCount = MAX_BATCH
    If NextCount < 1# Then NextCount = 1#
    m_BatchIterations = CLng(Fix(NextCount))
End Sub

Private Function AllocateExecutable(ByVal HexCode As String) As Long
    Dim Parts() As String, Bytes() As Byte, Count As Long, i As Long
    Dim Token As String, Address As Long, Previous As Long, Detail As String
    On Error GoTo Failed
    HexCode = Replace$(Replace$(Replace$(HexCode, vbCr, " "), vbLf, " "), vbTab, " ")
    Parts = Split(HexCode, " ")
    ReDim Bytes(0 To 65535)
    For i = 0 To UBound(Parts)
        Token = Parts(i)
        If Len(Token) <> 0 Then
            If Len(Token) <> 2 Then Err.Raise 5, , "Invalid embedded stress bytecode token."
            If Count >= 65536 Then Err.Raise 5, , "Embedded stress routine is too large."
            Bytes(Count) = CByte("&H" & Token)
            Count = Count + 1
        End If
    Next i
    If Count = 0 Then Err.Raise 5, , "Embedded stress routine is empty."
    Address = StressVirtualAlloc(0, Count, MEM_COMMIT_RESERVE, PAGE_READWRITE)
    If Address = 0 Then Err.Raise 7, , "Cannot allocate stress code memory."
    StressCopyMemory ByVal Address, Bytes(0), Count
    If StressVirtualProtect(Address, Count, PAGE_EXECUTE_READ, Previous) = 0 Then Err.Raise 5, , "Cannot protect executable stress code."
    If StressFlushInstructionCache(StressGetCurrentProcess(), Address, Count) = 0 Then Err.Raise 5, , "Cannot flush stress instruction cache."
    AllocateExecutable = Address
    Exit Function
Failed:
    Detail = Err.Description
    If Address <> 0 Then Call StressVirtualFree(Address, 0, MEM_RELEASE)
    Err.Raise vbObjectError + 2301, "StressLoop", Detail
End Function

Private Function CPUIDProbeHex() As String
    CPUIDProbeHex = "55 89 E5 53 57 8B 45 08 8B 7D 0C 31 C9 0F A2 89 07 89 5F 04 89 4F 08 89 57 0C 5F 5B 5D C2 10 00"
    'push ebp
    'mov ebp, esp
    'push ebx
    'push edi
    'mov eax, [ebp+8]       ; requested CPUID leaf
    'mov edi, [ebp+12]      ; output EAX, EBX, ECX, EDX
    'xor ecx, ecx          ; subleaf zero
    'cpuid
    'mov [edi], eax
    'mov [edi+4], ebx
    'mov [edi+8], ecx
    'mov [edi+12], edx
    'pop edi
    'pop ebx
    'pop ebp
    'ret 16
End Function

Private Function XGETBVProbeHex() As String
    XGETBVProbeHex = "55 89 E5 57 8B 7D 0C 31 C9 0F 01 D0 89 07 89 57 04 5F 5D C2 10 00"
    'push ebp
    'mov ebp, esp
    'push edi
    'mov edi, [ebp+12]      ; output low/high halves of XCR0
    'xor ecx, ecx          ; XCR0
    'xgetbv                ; guarded by XSAVE and OSXSAVE checks
    'mov [edi], eax
    'mov [edi+4], edx
    'pop edi
    'pop ebp
    'ret 16
End Function
